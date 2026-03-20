// 1. Importações e Configurações Iniciais
require('dotenv').config();
const express = require('express');
const mongoose = require('mongoose');
const cors = require('cors');

// --- IMPORTAÇÕES DA AUTOMAÇÃO E ARQUIVOS ---
const cron = require('node-cron');
const msal = require('@azure/msal-node');
const { Client } = require('@microsoft/microsoft-graph-client');
require('isomorphic-fetch');
const multer = require('multer');
const ExcelJS = require('exceljs'); // Nova biblioteca para preencher a planilha molde

const app = express();
app.use(cors({ origin: '*' })); 
app.use(express.json());

// Configuração do Recebedor de Arquivos (Limita em 15MB)
const upload = multer({ 
    storage: multer.memoryStorage(),
    limits: { fileSize: 15 * 1024 * 1024 } 
});

// 2. Conexão Segura com MongoDB
const mongoURI = process.env.MONGO_URI;

if (!mongoURI) {
    console.error("❌ ERRO FATAL: A variável de ambiente MONGO_URI não foi encontrada!");
    process.exit(1);
}

mongoose.connect(mongoURI)
    .then(() => {
        console.log('🟢 MongoDB Atlas Conectado com Segurança!');
        semearBanco(); 
    })
    .catch(err => console.error("❌ Erro ao conectar ao MongoDB:", err));

// ==========================================
// --- MODELOS DO BANCO DE DADOS ---
// ==========================================
const Apropriacao = mongoose.model('Apropriacao', new mongoose.Schema({ 
    data: String, 
    dados_dia: Object 
}));

const Funcionario = mongoose.model('Funcionario', new mongoose.Schema({ 
    mat: String, 
    nome: String,
    cargo: { type: String, default: 'Não informado' },
    custoDiario: { type: Number, default: 0 },
    isResidente: { type: Boolean, default: false },
    projetoResidente: { type: String, default: '' }
}));

const Projeto = mongoose.model('Projeto', new mongoose.Schema({ 
    codigo: String 
}));

const Documento = mongoose.model('Documento', new mongoose.Schema({
    nome: String,
    area: String,
    tipo: String,
    ext: String,
    tamanho: String,
    data: String,
    arquivoBase64: String 
}));

// --- FUNÇÃO DE SEMENTE (DATABASE SEED) ---
async function semearBanco() {
    const fCount = await Funcionario.countDocuments();
    if (fCount === 0) {
        const iniciais = [
            { mat: "79", nome: "Amarildo Fernandes Rosa", cargo: "Eletricista", custoDiario: 150 }, 
            { mat: "91", nome: "Diogo Bassi Rosa", cargo: "Encarregado", custoDiario: 250 }
        ];
        await Funcionario.insertMany(iniciais);
        await Projeto.insertMany([{codigo:'230304'}, {codigo:'242236'}, {codigo:'C522006'}, {codigo:'ATESTADO'}]);
        console.log('✅ Dados iniciais semeados no banco de dados!');
    }
}

// ==========================================
// 🚀 MOTOR DE AUTOMAÇÃO MICROSOFT 365
// ==========================================
const msalConfig = {
    auth: {
        clientId: process.env.CLIENT_ID,
        authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
        clientSecret: process.env.CLIENT_SECRET,
    }
};

const cca = new msal.ConfidentialClientApplication(msalConfig);

async function getGraphClient() {
    const authResponse = await cca.acquireTokenByClientCredential({ scopes: ['https://graph.microsoft.com/.default'] });
    return Client.init({ authProvider: (done) => done(null, authResponse.accessToken) });
}

function encodeShareUrl(url) {
    const base64 = Buffer.from(url).toString('base64');
    return 'u!' + base64.replace(/=/g, '').replace(/\//g, '_').replace(/\+/g, '-');
}

async function sincronizarPlanilha() {
    console.log("🔄 Iniciando sincronização com Microsoft 365...");
    try {
        if (!process.env.PLANILHA_URL) throw new Error("A variável PLANILHA_URL não foi configurada.");
        
        const client = await getGraphClient();
        const shareToken = encodeShareUrl(process.env.PLANILHA_URL);
        
        const driveItem = await client.api(`/shares/${shareToken}/driveItem`).get();
        const excel = await client.api(`/drives/${driveItem.parentReference.driveId}/items/${driveItem.id}/workbook/tables/Tabela1/rows`).get();
        
        let obrasAtualizadas = 0; let obrasRemovidas = 0;

        for (let row of excel.value) {
            const codigo = row.values[0][0]; 
            const status = row.values[0][1]; 

            if (!codigo) continue; 

            const codFormatado = codigo.toString().trim().toUpperCase();
            const statFormatado = status ? status.toString().trim().toLowerCase() : 'sim';

            if (statFormatado === 'não' || statFormatado === 'nao' || statFormatado === 'false') {
                await Projeto.deleteOne({ codigo: codFormatado });
                obrasRemovidas++;
            } else {
                await Projeto.findOneAndUpdate({ codigo: codFormatado }, { codigo: codFormatado }, { upsert: true });
                obrasAtualizadas++;
            }
        }
        console.log(`✅ Sincronização concluída! ${obrasAtualizadas} ativas | ${obrasRemovidas} inativas removidas.`);
        return { sucesso: true, obrasAtualizadas, obrasRemovidas };
    } catch (error) {
        console.error("❌ Erro ao ler Excel no 365:", error.message);
        throw error;
    }
}

// ==========================================
// 👷 AUTOMATIZADOR DE RESIDENTES
// ==========================================
async function lancarHorasResidentes() {
    console.log("🔄 Iniciando o salvamento automático de Residentes...");
    try {
        const objData = new Date();
        const tzOffset = objData.getTimezoneOffset() * 60000;
        const hoje = new Date(objData.getTime() - tzOffset);
        const dataStr = hoje.toISOString().split('T')[0];

        const diaSemana = hoje.getUTCDay();
        if (diaSemana === 0 || diaSemana === 6) {
            console.log("⏸️ Fim de semana: Lançamento de residentes pausado hoje.");
            return;
        }

        const residentes = await Funcionario.find({ isResidente: true });
        if (residentes.length === 0) return;

        let apropriacaoHoje = await Apropriacao.findOne({ data: dataStr });
        let dados_dia = apropriacaoHoje ? apropriacaoHoje.dados_dia : {};
        let teveMudanca = false;

        residentes.forEach(func => {
            if (!dados_dia[func.mat] && func.projetoResidente) {
                dados_dia[func.mat] = { h1: 9, p1: func.projetoResidente };
                teveMudanca = true;
            }
        });

        if (teveMudanca) {
            await Apropriacao.findOneAndUpdate(
                { data: dataStr },
                { dados_dia: dados_dia },
                { upsert: true }
            );
            console.log(`✅ Horas salvas automaticamente no banco para ${residentes.length} residentes!`);
        } else {
            console.log(`✅ Todos os residentes já estavam com as horas salvas hoje.`);
        }
    } catch (error) {
        console.error("❌ Erro ao lançar residentes:", error.message);
    }
}

// ⏰ O DESPERTADOR: Roda automaticamente todo dia às 02:00 da manhã
cron.schedule('0 2 * * *', async () => {
    await sincronizarPlanilha();
    await lancarHorasResidentes(); 
});


// ==========================================
// --- ROTAS DA API ---
// ==========================================

// --- ROTA INTELIGENTE: EXPORTAR PLANILHA FINANCEIRA DO BANCO ---
app.post('/api/financeiro/exportar', async (req, res) => {
    try {
        const { projeto, cabecalho, linhas } = req.body;
        
        // 1. Procura a planilha exata na aba de Documentos
        const templateDoc = await Documento.findOne({ nome: 'CONTROLE_FINANCEIRO_PROJETO.xlsx' });
        
        if (!templateDoc) {
            return res.status(404).json({ 
                erro: "Molde não encontrado! Vá na aba 'Documentos' e faça o upload do arquivo com o nome exato: CONTROLE_FINANCEIRO_PROJETO.xlsx" 
            });
        }

        // 2. Converte o Base64 de volta para Arquivo na memória
        const buffer = Buffer.from(templateDoc.arquivoBase64, 'base64');

        // 3. Abre o arquivo original preservando cores, fórmulas e estilos
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(buffer);
        const worksheet = workbook.getWorksheet(1);

        // 4. Injeta o Cabeçalho Financeiro
        worksheet.getCell('D3').value = cabecalho.empresa || 'CLIMATE';
        worksheet.getCell('D4').value = projeto;
        worksheet.getCell('C5').value = cabecalho.inicio ? cabecalho.inicio.split('-').reverse().join('/') : '';
        worksheet.getCell('E5').value = cabecalho.centroCusto;
        worksheet.getCell('C6').value = cabecalho.fim ? cabecalho.fim.split('-').reverse().join('/') : '';
        worksheet.getCell('E6').value = parseFloat(cabecalho.saldoBruto || 0);
        worksheet.getCell('C7').value = parseFloat(cabecalho.rateioCC || 0);
        worksheet.getCell('E7').value = parseFloat(cabecalho.saldoLiquido || 0);
        
        worksheet.getCell('E9').value = parseFloat(cabecalho.valorAntecipado || 0);
        
        worksheet.getCell('C10').value = parseFloat(cabecalho.receitaPrevista || 0);
        worksheet.getCell('E10').value = parseFloat(cabecalho.receitaReal || 0);
        worksheet.getCell('C11').value = parseFloat(cabecalho.impostoPrevisto || 0);
        worksheet.getCell('E11').value = parseFloat(cabecalho.impostoReal || 0);
        worksheet.getCell('C12').value = parseFloat(cabecalho.margemPrevista || 0);
        worksheet.getCell('E12').value = parseFloat(cabecalho.margemReal || 0);
        worksheet.getCell('C13').value = parseFloat(cabecalho.materialPrevisto || 0);
        worksheet.getCell('E13').value = parseFloat(cabecalho.materialReal || 0);
        worksheet.getCell('C14').value = parseFloat(cabecalho.servicoPrevisto || 0);
        worksheet.getCell('E14').value = parseFloat(cabecalho.servicoReal || 0);
        worksheet.getCell('C15').value = parseFloat(cabecalho.custoFinanPrevisto || 0);
        worksheet.getCell('E15').value = parseFloat(cabecalho.custoFinanReal || 0);

        // 5. Injeta os Lançamentos do Extrato (começando na linha 17)
        let row = 17;
        linhas.forEach(l => {
            worksheet.getCell(`B${row}`).value = l.data ? l.data.split('-').reverse().join('/') : '';
            worksheet.getCell(`C${row}`).value = l.tipo;
            worksheet.getCell(`D${row}`).value = l.desc;
            worksheet.getCell(`E${row}`).value = parseFloat(l.valorBruto || 0);
            worksheet.getCell(`F${row}`).value = parseFloat(l.valorLiquido || 0);
            row++;
        });

        // 6. Devolve o Excel pronto e impecável para o Front-end baixar
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="Financeiro_${projeto}.xlsx"`);
        
        await workbook.xlsx.write(res);
        res.end();
    } catch (err) {
        console.error("Erro na exportação via MongoDB:", err);
        res.status(500).json({ erro: err.message });
    }
});


// --- ROTAS DE DOCUMENTOS ---
app.post('/api/documentos', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) throw new Error("Nenhum arquivo enviado.");
        const { nome, area, ext, tamanho, data, tipo } = req.body;
        const arquivoBase64 = req.file.buffer.toString('base64'); 
        
        const novoDoc = new Documento({ nome, area, ext, tamanho, data, tipo, arquivoBase64 });
        await novoDoc.save();
        
        const docResumo = { ...novoDoc._doc };
        delete docResumo.arquivoBase64; 
        
        res.json({ sucesso: true, doc: docResumo });
    } catch (err) { res.status(500).json({ erro: err.message }); }
});

app.get('/api/documentos', async (req, res) => {
    try {
        const docs = await Documento.find().select('-arquivoBase64').sort({ _id: -1 });
        res.json(docs);
    } catch (err) { res.status(500).json({ erro: err.message }); }
});

app.get('/api/documentos/:id/download', async (req, res) => {
    try {
        const doc = await Documento.findById(req.params.id);
        res.json({ arquivoBase64: doc.arquivoBase64, tipo: doc.tipo, nome: doc.nome });
    } catch (err) { res.status(500).json({ erro: err.message }); }
});

app.delete('/api/documentos/:id', async (req, res) => {
    try {
        await Documento.findByIdAndDelete(req.params.id);
        res.json({ sucesso: true });
    } catch (err) { res.status(500).json({ erro: err.message }); }
});

// --- ROTAS DE PROJETOS E SINCRONIZAÇÃO ---
app.get('/api/projetos/force-sync', async (req, res) => {
    try {
        const resultado = await sincronizarPlanilha();
        res.json(resultado);
    } catch (err) { res.status(500).json({ erro: err.message }); }
});

app.get('/api/projetos', async (req, res) => {
    const p = await Projeto.find().sort({ codigo: 1 });
    res.json(p.map(x => x.codigo));
});

app.post('/api/projetos', async (req, res) => { 
    try {
        await new Projeto({codigo: req.body.codigo}).save(); 
        res.json({sucesso:true}); 
    } catch (err) { res.status(500).json({ erro: err.message }); }
});

app.delete('/api/projetos/:cod', async (req, res) => { 
    await Projeto.deleteOne({ codigo: req.params.cod }); 
    res.json({sucesso:true}); 
});

// --- ROTAS DE APROPRIAÇÃO ---
app.get('/api/apropriacao', async (req, res) => {
    try {
        const todos = await Apropriacao.find();
        let banco = {};
        todos.forEach(reg => { banco[reg.data] = reg.dados_dia; });
        res.json(banco);
    } catch (err) { res.status(500).json({ erro: err.message }); }
});

app.post('/api/apropriacao', async (req, res) => {
    try {
        await Apropriacao.findOneAndUpdate({ data: req.body.data }, { dados_dia: req.body.dados_dia }, { upsert: true });
        res.json({ sucesso: true });
    } catch (err) { res.status(500).json({ erro: err.message }); }
});

// --- ROTAS DE EQUIPE ---
app.get('/api/equipe', async (req, res) => res.json(await Funcionario.find().sort({ nome: 1 })));

app.post('/api/equipe', async (req, res) => { 
    try {
        await Funcionario.findOneAndUpdate({ mat: req.body.mat }, req.body, { upsert: true }); 
        res.json({sucesso:true}); 
    } catch (err) { res.status(500).json({ erro: err.message }); }
});

app.delete('/api/equipe/:mat', async (req, res) => { 
    await Funcionario.deleteOne({ mat: req.params.mat }); 
    res.json({sucesso:true}); 
});

// --- ROTA DE LOGIN ---
app.post('/api/login', (req, res) => {
    const { usuario, senha } = req.body;
    if (usuario === 'gerencia' && senha === 'L2pgerencia2026!') res.json({ tipo: 'admin', nome: 'Gerência' });
    else if (usuario === 'analista' && senha === 'L2panalista2026') res.json({ tipo: 'user', nome: 'Analista L2P' });
    else res.status(401).json({ erro: 'Usuário ou senha incorretos' });
});

// 3. Inicialização do Servidor
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`🚀 Servidor L2P rodando na porta ${PORT}`));
