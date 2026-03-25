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
const ExcelJS = require('exceljs'); // IMPORTAÇÃO DO GERADOR DE EXCEL

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

// 🔥 NOVO MODELO: Guarda as planilhas separadas de cada projeto
const Financeiro = mongoose.model('Financeiro', new mongoose.Schema({ 
    projeto: String, 
    cabecalho: Object,
    linhas: Array
}));

const ItemEstoque = mongoose.model('Estoque', new mongoose.Schema({
    codigo: String,
    descricao: String,
    unidade: { type: String, default: 'Un' },
    quantidade: { type: Number, default: 0 },
    estoqueMinimo: { type: Number, default: 0 },
    categoria: String,
    ultimaAtualizacao: String
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
        if (diaSemana === 0 || diaSemana === 6) return;

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
            await Apropriacao.findOneAndUpdate({ data: dataStr }, { dados_dia: dados_dia }, { upsert: true });
            console.log(`✅ Horas salvas automaticamente para ${residentes.length} residentes!`);
        }
    } catch (error) { console.error("❌ Erro ao lançar residentes:", error.message); }
}

cron.schedule('0 2 * * *', async () => {
    await sincronizarPlanilha();
    await lancarHorasResidentes(); 
});


// ==========================================
// --- ROTAS DA API ---
// ==========================================

// 🔥 NOVAS ROTAS DO FINANCEIRO (VISUALIZAÇÃO NATIVA NO SITE)
app.get('/api/financeiro/:projeto', async (req, res) => {
    try {
        const fin = await Financeiro.findOne({ projeto: req.params.projeto });
        res.json(fin || { cabecalho: {}, linhas: [] });
    } catch (err) { res.status(500).json({ erro: err.message }); }
});

app.post('/api/financeiro/:projeto', async (req, res) => {
    try {
        await Financeiro.findOneAndUpdate(
            { projeto: req.params.projeto }, 
            { cabecalho: req.body.cabecalho, linhas: req.body.linhas }, 
            { upsert: true }
        );
        res.json({ sucesso: true });
    } catch (err) { res.status(500).json({ erro: err.message }); }
});

// 🔥 ROTA DE EXPORTAÇÃO (GERA O EXCEL REAL E BAIXA)
app.post('/api/financeiro/exportar', async (req, res) => {
    try {
        const { projeto, cabecalho, linhas } = req.body;
        
        const templateDoc = await Documento.findOne({ nome: 'CONTROLE_FINANCEIRO_PROJETO.xlsx' });
        if (!templateDoc) return res.status(404).json({ erro: "Molde não encontrado! Suba o arquivo CONTROLE_FINANCEIRO_PROJETO.xlsx na aba de Documentos." });

        const buffer = Buffer.from(templateDoc.arquivoBase64, 'base64');
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(buffer);
        const worksheet = workbook.getWorksheet('Planilha1 (2)') || workbook.worksheets[0];
        
        if (!worksheet) throw new Error("Aba não encontrada.");

        // Injeta os valores nas células exatas da L2P
        const val = (v) => v ? parseFloat(v) : 0;
        
        worksheet.getCell('D4').value = projeto; // Obra
        worksheet.getCell('D5').value = cabecalho.centroCusto || projeto;
        if (cabecalho.inicio) worksheet.getCell('C5').value = cabecalho.inicio.split('-').reverse().join('/');
        if (cabecalho.fim) worksheet.getCell('C6').value = cabecalho.fim.split('-').reverse().join('/');
        
        // Mão de Obra
        worksheet.getCell('C14').value = val(cabecalho.servicoReal); 
        worksheet.getCell('E14').value = val(cabecalho.servicoReal); 

        // Adiciona as Linhas (Lançamentos manuais feitos no site)
        let row = 17;
        linhas.forEach(l => {
            const currentRow = worksheet.getRow(row);
            if (l.data) currentRow.getCell('B').value = l.data.split('-').reverse().join('/');
            if (l.tipo) currentRow.getCell('C').value = l.tipo;
            if (l.desc) currentRow.getCell('D').value = l.desc;
            if (l.valorBruto) currentRow.getCell('E').value = parseFloat(l.valorBruto);
            if (l.valorLiquido) currentRow.getCell('F').value = parseFloat(l.valorLiquido);
            row++;
        });

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="Financeiro_${projeto}.xlsx"`);
        await workbook.xlsx.write(res);
        res.end();
    } catch (err) {
        console.error("❌ ERRO NA EXPORTAÇÃO:", err.message);
        res.status(500).json({ erro: err.message });
    }
});


// --- ROTAS RESTANTES (Documentos, Projetos, Equipe, Apropriação, Estoque) ---
app.post('/api/documentos', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) throw new Error("Nenhum arquivo enviado.");
        const { nome, area, ext, tamanho, data, tipo } = req.body;
        const novoDoc = new Documento({ nome, area, ext, tamanho, data, tipo, arquivoBase64: req.file.buffer.toString('base64') });
        await novoDoc.save();
        const docResumo = { ...novoDoc._doc }; delete docResumo.arquivoBase64; 
        res.json({ sucesso: true, doc: docResumo });
    } catch (err) { res.status(500).json({ erro: err.message }); }
});
app.get('/api/documentos', async (req, res) => { res.json(await Documento.find().select('-arquivoBase64').sort({ _id: -1 })); });
app.get('/api/documentos/:id/download', async (req, res) => { res.json(await Documento.findById(req.params.id)); });
app.delete('/api/documentos/:id', async (req, res) => { await Documento.findByIdAndDelete(req.params.id); res.json({ sucesso: true }); });

app.get('/api/projetos/force-sync', async (req, res) => { res.json(await sincronizarPlanilha()); });
app.get('/api/projetos', async (req, res) => { const p = await Projeto.find().sort({ codigo: 1 }); res.json(p.map(x => x.codigo)); });
app.post('/api/projetos', async (req, res) => { await new Projeto({codigo: req.body.codigo}).save(); res.json({sucesso:true}); });
app.delete('/api/projetos/:cod', async (req, res) => { await Projeto.deleteOne({ codigo: req.params.cod }); res.json({sucesso:true}); });

app.get('/api/apropriacao', async (req, res) => {
    const todos = await Apropriacao.find(); let banco = {};
    todos.forEach(reg => { banco[reg.data] = reg.dados_dia; }); res.json(banco);
});
app.post('/api/apropriacao', async (req, res) => { await Apropriacao.findOneAndUpdate({ data: req.body.data }, { dados_dia: req.body.dados_dia }, { upsert: true }); res.json({ sucesso: true }); });

app.get('/api/equipe', async (req, res) => res.json(await Funcionario.find().sort({ nome: 1 })));
app.post('/api/equipe', async (req, res) => { await Funcionario.findOneAndUpdate({ mat: req.body.mat }, req.body, { upsert: true }); res.json({sucesso:true}); });
app.delete('/api/equipe/:mat', async (req, res) => { await Funcionario.deleteOne({ mat: req.params.mat }); res.json({sucesso:true}); });

app.get('/api/estoque', async (req, res) => res.json(await ItemEstoque.find().sort({ descricao: 1 })));
app.post('/api/estoque', async (req, res) => { const dados = req.body; dados.ultimaAtualizacao = new Date().toLocaleString('pt-BR'); await new ItemEstoque(dados).save(); res.json({ sucesso: true }); });
app.put('/api/estoque/:id', async (req, res) => { const dados = req.body; dados.ultimaAtualizacao = new Date().toLocaleString('pt-BR'); await ItemEstoque.findByIdAndUpdate(req.params.id, dados); res.json({ sucesso: true }); });
app.delete('/api/estoque/:id', async (req, res) => { await ItemEstoque.findByIdAndDelete(req.params.id); res.json({ sucesso: true }); });

app.post('/api/login', (req, res) => {
    const { usuario, senha } = req.body;
    if (usuario === 'gerencia' && senha === 'L2pgerencia2026!') res.json({ tipo: 'admin', nome: 'Gerência' });
    else if (usuario === 'analista' && senha === 'L2panalista2026') res.json({ tipo: 'user', nome: 'Analista L2P' });
    else res.status(401).json({ erro: 'Usuário ou senha incorretos' });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`🚀 Servidor L2P rodando na porta ${PORT}`));
