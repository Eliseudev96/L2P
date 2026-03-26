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

const ItemEstoque = mongoose.model('Estoque', new mongoose.Schema({
    codigo: String,
    descricao: String,
    unidade: { type: String, default: 'Un' },
    quantidade: { type: Number, default: 0 },
    estoqueMinimo: { type: Number, default: 0 },
    categoria: String,
    ultimaAtualizacao: String
}));

// --- MODELOS DE SEGURANÇA E CONFIGURAÇÕES ---
const Usuario = mongoose.model('Usuario', new mongoose.Schema({
    nome: String,
    email: String,
    cargo: String,
    nivel: String,
    status: String,
    senha: String // Em produção futura, aplicar hash (bcrypt)
}));

const Empresa = mongoose.model('Empresa', new mongoose.Schema({
    razaoSocial: String, cnpj: String, inscricao: String, endereco: String
}));

const Notificacao = mongoose.model('Notificacao', new mongoose.Schema({
    docNovo: Boolean, horaExtra: Boolean, relatorioFin: Boolean
}));

const Auditoria = mongoose.model('Auditoria', new mongoose.Schema({
    data: String, usuario: String, acao: String, ip: String
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

// --- ROTAS DE USUÁRIOS E SEGURANÇA ---
app.get('/api/usuarios', async (req, res) => {
    try {
        const users = await Usuario.find().select('-senha'); 
        res.json(users);
    } catch (err) { res.status(500).json({ erro: err.message }); }
});

app.post('/api/usuarios', async (req, res) => {
    try {
        const novoUser = new Usuario(req.body);
        await novoUser.save();
        res.json(novoUser);
    } catch (err) { res.status(500).json({ erro: err.message }); }
});

app.put('/api/usuarios/:id', async (req, res) => {
    try {
        const updateData = { ...req.body };
        if (!updateData.senha) delete updateData.senha; 
        
        await Usuario.findByIdAndUpdate(req.params.id, updateData);
        res.json({ sucesso: true });
    } catch (err) { res.status(500).json({ erro: err.message }); }
});

app.delete('/api/usuarios/:id', async (req, res) => {
    try {
        await Usuario.findByIdAndDelete(req.params.id);
        res.json({ sucesso: true });
    } catch (err) { res.status(500).json({ erro: err.message }); }
});

// --- ROTAS DA EMPRESA ---
app.get('/api/empresa', async (req, res) => {
    let emp = await Empresa.findOne();
    if (!emp) emp = await Empresa.create({ razaoSocial: '', cnpj: '', inscricao: '', endereco: '' });
    res.json(emp);
});

app.put('/api/empresa', async (req, res) => {
    let emp = await Empresa.findOne();
    if (emp) { Object.assign(emp, req.body); await emp.save(); } 
    else { emp = await Empresa.create(req.body); }
    res.json(emp);
});

// --- ROTAS DE NOTIFICAÇÕES ---
app.get('/api/notificacoes', async (req, res) => {
    let notif = await Notificacao.findOne();
    if (!notif) notif = await Notificacao.create({ docNovo: false, horaExtra: false, relatorioFin: false });
    res.json(notif);
});

app.put('/api/notificacoes', async (req, res) => {
    let notif = await Notificacao.findOne();
    if (notif) { Object.assign(notif, req.body); await notif.save(); } 
    else { notif = await Notificacao.create(req.body); }
    res.json(notif);
});

// --- ROTA DE AUDITORIA ---
app.get('/api/auditoria', async (req, res) => {
    try {
        const logs = await Auditoria.find().sort({ _id: -1 }).limit(50);
        res.json(logs);
    } catch (err) { res.status(500).json({ erro: err.message }); }
});

// --- ROTA DE LOGIN (Atualizada) ---
app.post('/api/login', async (req, res) => {
    const { usuario, senha } = req.body;
    
    try {
        // Verifica no banco de dados primeiro
        const userDb = await Usuario.findOne({ email: usuario, senha: senha, status: 'Ativo' });
        
        if (userDb) {
            return res.json({ tipo: userDb.nivel, nome: userDb.nome });
        }
        
        // Fallback para login de admin mestre hardcoded caso precisem entrar
        if (usuario === 'gerencia' && senha === 'L2pgerencia2026!') {
            return res.json({ tipo: 'Admin', nome: 'Gerência' });
        }
        
        res.status(401).json({ erro: 'Usuário ou senha incorretos, ou cadastro inativo.' });
    } catch (err) {
        res.status(500).json({ erro: err.message });
    }
});


// 🔥 ROTA DE DOCUMENTOS FINANCEIROS (GOOGLE DRIVE)
app.post('/api/financeiro/salvar-documento', async (req, res) => {
    try {
        const { sheetId, projeto } = req.body;
        if (!sheetId || !projeto) throw new Error("Faltam dados da planilha.");

        const exportUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=xlsx`;
        
        const fetch = require('isomorphic-fetch');
        const response = await fetch(exportUrl);
        
        if (!response.ok) {
            throw new Error(`O Google bloqueou o download. Status: ${response.status}`);
        }

        const contentType = response.headers.get('content-type');
        if (contentType && contentType.includes('text/html')) {
            throw new Error("Acesso negado pelo Google. Verifique as permissões do Apps Script.");
        }
        
        let buffer;
        if (typeof response.buffer === 'function') {
            buffer = await response.buffer(); 
        } else {
            const arrayBuffer = await response.arrayBuffer(); 
            buffer = Buffer.from(arrayBuffer);
        }
        
        const arquivoBase64 = buffer.toString('base64');
        const dataAtual = new Date().toISOString().split('T')[0];
        const tamanhoKB = (buffer.length / 1024).toFixed(2) + ' KB';
        const nomeArquivo = `Financeiro_Obra_${projeto}_${dataAtual}`;

        const novoDoc = new Documento({
            nome: nomeArquivo,
            area: 'Financeiro',
            tipo: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            ext: 'xlsx',
            tamanho: tamanhoKB,
            data: dataAtual,
            arquivoBase64: arquivoBase64
        });

        await novoDoc.save();
        res.json({ sucesso: true });
    } catch (err) {
        console.error("❌ Erro ao salvar documento:", err.message);
        res.status(500).json({ erro: err.message });
    }
});

// PONTE SEGURA PARA O GOOGLE APPS SCRIPT
app.get('/api/google-proxy', async (req, res) => {
    try {
        const fetch = require('isomorphic-fetch');
        const url = req.query.url;
        const response = await fetch(url);
        
        const text = await response.text(); 
        
        try {
            const data = JSON.parse(text);
            res.json(data);
        } catch (e) {
            console.error("❌ Google devolveu HTML (Bloqueio de Permissão):", text.substring(0, 200));
            res.status(500).json({ erro: "Bloqueio de segurança do Google." });
        }
    } catch (err) {
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

// --- ROTAS DE ESTOQUE ---
app.get('/api/estoque', async (req, res) => {
    try {
        const itens = await ItemEstoque.find().sort({ descricao: 1 });
        res.json(itens);
    } catch (err) { res.status(500).json({ erro: err.message }); }
});

app.post('/api/estoque', async (req, res) => {
    try {
        const dados = req.body;
        dados.ultimaAtualizacao = new Date().toLocaleString('pt-BR');
        const item = new ItemEstoque(dados);
        await item.save();
        res.json({ sucesso: true });
    } catch (err) { res.status(500).json({ erro: err.message }); }
});

app.put('/api/estoque/:id', async (req, res) => {
    try {
        const dados = req.body;
        dados.ultimaAtualizacao = new Date().toLocaleString('pt-BR');
        await ItemEstoque.findByIdAndUpdate(req.params.id, dados);
        res.json({ sucesso: true });
    } catch (err) { res.status(500).json({ erro: err.message }); }
});

app.delete('/api/estoque/:id', async (req, res) => {
    try {
        await ItemEstoque.findByIdAndDelete(req.params.id);
        res.json({ sucesso: true });
    } catch (err) { res.status(500).json({ erro: err.message }); }
});

// 3. Inicialização do Servidor
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`🚀 Servidor L2P rodando na porta ${PORT}`));
