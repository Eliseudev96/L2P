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

const Alerta = mongoose.model('Alerta', new mongoose.Schema({ 
    data: String, 
    texto: String, 
    tipo: String, 
    lido: { type: Boolean, default: false } 
}));

// 📅 MODELO DA AGENDA COMPARTILHADA
const Evento = mongoose.model('Evento', new mongoose.Schema({
    titulo: String,
    dataInicio: String, // formato YYYY-MM-DD ou ISO
    dataFim: String,
    tipo: String,       // ex: 'reuniao', 'manutencao', 'obra', etc
    descricao: String,
    responsavel: String
}));

// 💰 MODELO DO CONTROLE FINANCEIRO NATIVO (MONGODB)
const LancamentoFinanceiro = mongoose.model('LancamentoFinanceiro', new mongoose.Schema({
    projeto: String, // Código da Obra
    data: String,    // YYYY-MM-DD
    tipo: String,    // 'Despesa' ou 'Receita'
    categoria: String, // 'Material', 'Alimentação', 'Terceiros', 'Faturamento', etc.
    descricao: String,
    valor: Number
}));

// ==========================================
// 🕵️‍♂️ FUNÇÃO DE AUDITORIA E ALERTAS
// ==========================================
async function registrarLog(req, acao, nomeForcado = null) {
    try {
        const ip = (req.headers && req.headers['x-forwarded-for']) || (req.socket && req.socket.remoteAddress) || 'IP Desconhecido';
        const dataHora = new Date().toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' });
        const usuarioLogado = nomeForcado || (req.headers && req.headers['x-usuario']) || 'Usuário Não Identificado';
        await Auditoria.create({ data: dataHora, usuario: usuarioLogado, acao: acao, ip: ip });
    } catch (e) {
        console.error("⚠️ Falha ao registrar log de auditoria:", e.message);
    }
}

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
        
        await registrarLog({ headers: {}, socket: { remoteAddress: 'Servidor' } }, "Sincronização de obras com MS365 concluída.", "Robô da Madrugada");
        
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
            
            await registrarLog({ headers: {}, socket: { remoteAddress: 'Servidor' } }, `Lançou 9h automáticas para ${residentes.length} residentes.`, "Robô da Madrugada");
            
            console.log(`✅ Horas salvas automaticamente no banco para ${residentes.length} residentes!`);
        } else {
            console.log(`✅ Todos os residentes já estavam com as horas salvas hoje.`);
        }
    } catch (error) {
        console.error("❌ Erro ao lançar residentes:", error.message);
    }
}

cron.schedule('0 2 * * *', async () => {
    await sincronizarPlanilha();
    await lancarHorasResidentes(); 
});


// ==========================================
// --- ROTAS DA API ---
// ==========================================

// --- ROTA DE LOGIN REAL ---
app.post('/api/login', async (req, res) => {
    const { usuario, senha } = req.body;
    try {
        const userDb = await Usuario.findOne({ email: usuario, senha: senha, status: 'Ativo' });
        
        if (userDb) {
            await registrarLog(req, 'Entrou no sistema', userDb.nome);
            return res.json({ tipo: userDb.nivel, nome: userDb.nome });
        }
        
        if (usuario === 'gerencia' && senha === 'L2pgerencia2026!') {
            await registrarLog(req, 'Acessou o sistema via Chave Mestra', 'Gerência');
            return res.json({ tipo: 'Admin', nome: 'Gerência' });
        }
        
        await registrarLog(req, `Tentativa falha de login (E-mail: ${usuario})`, 'Desconhecido');
        res.status(401).json({ erro: 'Usuário ou senha incorretos' });
    } catch (err) { res.status(500).json({ erro: err.message }); }
});

// --- ROTAS DE SEGURANÇA E CONFIGURAÇÕES ---
app.get('/api/usuarios', async (req, res) => {
    try { res.json(await Usuario.find().select('-senha')); } 
    catch (err) { res.status(500).json({ erro: err.message }); }
});

app.post('/api/usuarios', async (req, res) => {
    try {
        const novoUser = new Usuario(req.body);
        await novoUser.save();
        await registrarLog(req, `Criou novo utilizador: ${novoUser.nome}`);
        res.json(novoUser);
    } catch (err) { res.status(500).json({ erro: err.message }); }
});

app.put('/api/usuarios/:id', async (req, res) => {
    try {
        const updateData = { ...req.body };
        if (!updateData.senha) delete updateData.senha; 
        await Usuario.findByIdAndUpdate(req.params.id, updateData);
        await registrarLog(req, `Editou o perfil do utilizador: ${updateData.nome}`);
        res.json({ sucesso: true });
    } catch (err) { res.status(500).json({ erro: err.message }); }
});

app.delete('/api/usuarios/:id', async (req, res) => {
    try {
        const user = await Usuario.findById(req.params.id);
        await Usuario.findByIdAndDelete(req.params.id);
        await registrarLog(req, `Excluiu o utilizador: ${user?.nome}`);
        res.json({ sucesso: true });
    } catch (err) { res.status(500).json({ erro: err.message }); }
});

app.get('/api/empresa', async (req, res) => {
    let emp = await Empresa.findOne();
    if (!emp) emp = await Empresa.create({ razaoSocial: '', cnpj: '', inscricao: '', endereco: '' });
    res.json(emp);
});

app.put('/api/empresa', async (req, res) => {
    let emp = await Empresa.findOne();
    if (emp) { Object.assign(emp, req.body); await emp.save(); } 
    else { emp = await Empresa.create(req.body); }
    await registrarLog(req, `Alterou os dados oficiais da Empresa`);
    res.json(emp);
});

app.get('/api/notificacoes', async (req, res) => {
    let notif = await Notificacao.findOne();
    if (!notif) notif = await Notificacao.create({ docNovo: false, horaExtra: false, relatorioFin: false });
    res.json(notif);
});

app.put('/api/notificacoes', async (req, res) => {
    let notif = await Notificacao.findOne();
    if (notif) { Object.assign(notif, req.body); await notif.save(); } 
    else { notif = await Notificacao.create(req.body); }
    await registrarLog(req, `Alterou as regras de Notificação`);
    res.json(notif);
});

app.get('/api/auditoria', async (req, res) => {
    try {
        const logs = await Auditoria.find().sort({ _id: -1 }).limit(100);
        res.json(logs);
    } catch (err) { res.status(500).json({ erro: err.message }); }
});

// --- ROTAS DE DOCUMENTOS ---
app.post('/api/documentos', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) throw new Error("Nenhum arquivo enviado.");
        const { nome, area, ext, tamanho, data, tipo } = req.body;
        const arquivoBase64 = req.file.buffer.toString('base64'); 
        
        const novoDoc = new Documento({ nome, area, ext, tamanho, data, tipo, arquivoBase64 });
        await novoDoc.save();
        
        await registrarLog(req, `Fez upload do documento: ${nome} na área ${area}`);
        
        const configAlertas = await Notificacao.findOne();
        if (configAlertas && configAlertas.docNovo) {
            const dataHora = new Date().toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' });
            await Alerta.create({ 
                data: dataHora, 
                texto: `Novo documento anexado na área ${area}: ${nome}`, 
                tipo: 'sucesso' 
            });
        }
        
        const docResumo = { ...novoDoc._doc }; delete docResumo.arquivoBase64; 
        res.json({ sucesso: true, doc: docResumo });
    } catch (err) { res.status(500).json({ erro: err.message }); }
});

app.get('/api/documentos', async (req, res) => {
    try { res.json(await Documento.find().select('-arquivoBase64').sort({ _id: -1 })); } 
    catch (err) { res.status(500).json({ erro: err.message }); }
});

app.get('/api/documentos/:id/download', async (req, res) => {
    try {
        const doc = await Documento.findById(req.params.id);
        await registrarLog(req, `Fez download do documento: ${doc.nome}`);
        res.json({ arquivoBase64: doc.arquivoBase64, tipo: doc.tipo, nome: doc.nome });
    } catch (err) { res.status(500).json({ erro: err.message }); }
});

app.delete('/api/documentos/:id', async (req, res) => {
    try {
        const doc = await Documento.findById(req.params.id);
        await Documento.findByIdAndDelete(req.params.id);
        await registrarLog(req, `Excluiu o documento: ${doc?.nome}`);
        res.json({ sucesso: true });
    } catch (err) { res.status(500).json({ erro: err.message }); }
});

// --- ROTAS DE PROJETOS ---
app.get('/api/projetos/force-sync', async (req, res) => {
    try {
        const resultado = await sincronizarPlanilha();
        await registrarLog(req, `Forçou a sincronização de obras com o M365 manualmente`);
        res.json(resultado);
    } catch (err) { res.status(500).json({ erro: err.message }); }
});

app.get('/api/projetos', async (req, res) => {
    const p = await Projeto.find().sort({ codigo: 1 }); res.json(p.map(x => x.codigo));
});

app.post('/api/projetos', async (req, res) => { 
    try { await new Projeto({codigo: req.body.codigo}).save(); res.json({sucesso:true}); } 
    catch (err) { res.status(500).json({ erro: err.message }); }
});

app.delete('/api/projetos/:cod', async (req, res) => { 
    await Projeto.deleteOne({ codigo: req.params.cod }); res.json({sucesso:true}); 
});

// --- ROTAS DE APROPRIAÇÃO ---
app.get('/api/apropriacao', async (req, res) => {
    try {
        const todos = await Apropriacao.find();
        let banco = {}; todos.forEach(reg => { banco[reg.data] = reg.dados_dia; });
        res.json(banco);
    } catch (err) { res.status(500).json({ erro: err.message }); }
});

app.post('/api/apropriacao', async (req, res) => {
    try {
        await Apropriacao.findOneAndUpdate({ data: req.body.data }, { dados_dia: req.body.dados_dia }, { upsert: true });
        await registrarLog(req, `Salvou apropriacao de horas referente ao dia ${req.body.data}`);
        
        const configAlertas = await Notificacao.findOne();
        if (configAlertas && configAlertas.horaExtra) {
            const dados = req.body.dados_dia;
            for (const mat in dados) {
                const totalH = (parseFloat(dados[mat].h1) || 0) + (parseFloat(dados[mat].h2) || 0) + (parseFloat(dados[mat].h3) || 0);
                if (totalH > 10) {
                    const func = await Funcionario.findOne({ mat });
                    const nomeFunc = func ? func.nome : `Matrícula ${mat}`;
                    const dataHora = new Date().toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' });
                    
                    const alertaExistente = await Alerta.findOne({ 
                        texto: { $regex: new RegExp(`excedeu 10h no dia ${req.body.data}`) },
                        texto: { $regex: new RegExp(nomeFunc) }
                    });
                    
                    if (!alertaExistente) {
                        await Alerta.create({ 
                            data: dataHora, 
                            texto: `Atenção: ${nomeFunc} excedeu 10h no dia ${req.body.data} (${totalH}h totais)`, 
                            tipo: 'alerta' 
                        });
                    }
                }
            }
        }
        
        res.json({ sucesso: true });
    } catch (err) { res.status(500).json({ erro: err.message }); }
});

// --- ROTAS DE EQUIPE ---
app.get('/api/equipe', async (req, res) => res.json(await Funcionario.find().sort({ nome: 1 })));

app.post('/api/equipe', async (req, res) => { 
    try {
        await Funcionario.findOneAndUpdate({ mat: req.body.mat }, req.body, { upsert: true }); 
        await registrarLog(req, `Adicionou/Editou o funcionário na Tabela de Custos: ${req.body.nome}`);
        res.json({sucesso:true}); 
    } catch (err) { res.status(500).json({ erro: err.message }); }
});

app.delete('/api/equipe/:mat', async (req, res) => { 
    await Funcionario.deleteOne({ mat: req.params.mat }); 
    await registrarLog(req, `Excluiu um funcionário da Tabela de Custos (Matrícula: ${req.params.mat})`);
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
        await registrarLog(req, `Cadastrou o item ${dados.descricao} no Estoque`);
        res.json({ sucesso: true });
    } catch (err) { res.status(500).json({ erro: err.message }); }
});

app.put('/api/estoque/:id', async (req, res) => {
    try {
        const dados = req.body;
        dados.ultimaAtualizacao = new Date().toLocaleString('pt-BR');
        await ItemEstoque.findByIdAndUpdate(req.params.id, dados);
        await registrarLog(req, `Atualizou as informações de um item no Estoque`);
        res.json({ sucesso: true });
    } catch (err) { res.status(500).json({ erro: err.message }); }
});

app.delete('/api/estoque/:id', async (req, res) => {
    try {
        const item = await ItemEstoque.findById(req.params.id);
        await ItemEstoque.findByIdAndDelete(req.params.id);
        await registrarLog(req, `Excluiu o item ${item?.descricao} do Estoque`);
        res.json({ sucesso: true });
    } catch (err) { res.status(500).json({ erro: err.message }); }
});

// --- ROTAS DE ALERTAS PARA O FRONTEND ---
app.get('/api/alertas', async (req, res) => {
    try { res.json(await Alerta.find().sort({ _id: -1 }).limit(30)); } 
    catch (err) { res.status(500).json({ erro: err.message }); }
});

app.post('/api/alertas/marcar-lidos', async (req, res) => {
    try { await Alerta.updateMany({ lido: false }, { $set: { lido: true } }); res.json({ sucesso: true }); } 
    catch (err) { res.status(500).json({ erro: err.message }); }
});

app.delete('/api/alertas', async (req, res) => {
    try { await Alerta.deleteMany({}); res.json({ sucesso: true }); } 
    catch (err) { res.status(500).json({ erro: err.message }); }
});

// --- ROTAS DA AGENDA COMPARTILHADA ---
app.get('/api/eventos', async (req, res) => {
    try {
        const eventos = await Evento.find().sort({ dataInicio: 1 });
        res.json(eventos);
    } catch (err) { res.status(500).json({ erro: err.message }); }
});

app.post('/api/eventos', async (req, res) => {
    try {
        const novoEvento = new Evento(req.body);
        await novoEvento.save();
        
        const usuarioLogado = req.headers['x-usuario'] || 'Usuário Desconhecido';
        await registrarLog(req, `Agendou compromisso: ${novoEvento.titulo}`);

        if (usuarioLogado !== 'Gerência' && usuarioLogado !== 'Gerência L2P') {
            const dataHora = new Date().toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' });
            await Alerta.create({
                data: dataHora,
                texto: `${usuarioLogado} agendou: ${novoEvento.titulo} (${novoEvento.dataInicio})`,
                tipo: 'alerta'
            });
        }

        res.json({ sucesso: true, evento: novoEvento });
    } catch (err) { res.status(500).json({ erro: err.message }); }
});

app.put('/api/eventos/:id', async (req, res) => {
    try {
        await Evento.findByIdAndUpdate(req.params.id, req.body);
        await registrarLog(req, `Alterou o evento da agenda: ${req.body.titulo}`);
        res.json({ sucesso: true });
    } catch (err) { res.status(500).json({ erro: err.message }); }
});

app.delete('/api/eventos/:id', async (req, res) => {
    try {
        const evento = await Evento.findById(req.params.id);
        await Evento.findByIdAndDelete(req.params.id);
        
        const usuarioLogado = req.headers['x-usuario'] || 'Usuário Desconhecido';
        await registrarLog(req, `Cancelou evento da agenda: ${evento?.titulo}`);

        if (usuarioLogado !== 'Gerência' && usuarioLogado !== 'Gerência L2P' && evento) {
            const dataHora = new Date().toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' });
            await Alerta.create({
                data: dataHora,
                texto: `${usuarioLogado} CANCELOU o evento: ${evento.titulo}`,
                tipo: 'urgente'
            });
        }

        res.json({ sucesso: true });
    } catch (err) { res.status(500).json({ erro: err.message }); }
});


// ==========================================
// 💰 INTEGRAÇÃO NATIVA: LER E SALVAR NO SHAREPOINT (l2pengenharialtda)
// ==========================================

// Função à prova de falhas para buscar o arquivo no SharePoint
// ==========================================
// 💰 SHAREPOINT DINÂMICO (BUSCAR / LER / EDITAR)
// ==========================================

// 🔎 Pega site + drive automaticamente
async function getDriveComercial(client) {
    const site = await client.api('/sites/l2pengenharialtda.sharepoint.com:/sites/Servidor').get();
    const drives = await client.api(`/sites/${site.id}/drives`).get();

    const drive = drives.value.find(d => d.name === 'Comercial') || drives.value[0];

    return {
        driveId: drive.id
    };
}

// 🔎 1. BUSCAR PLANILHAS
app.get('/api/planilhas/search', async (req, res) => {
    try {
        const { q } = req.query;
        const client = await getGraphClient();

        const { driveId } = await getDriveComercial(client);

        const result = await client
            .api(`/drives/${driveId}/root/search(q='${q}')`)
            .get();

        const arquivos = result.value
            .filter(f => f.name.endsWith('.xlsx'))
            .map(f => ({
                id: f.id,
                nome: f.name,
                driveId
            }));

        res.json(arquivos);

    } catch (err) {
        console.error("❌ Erro ao buscar planilhas:", err.message);
        res.status(500).json({ erro: err.message });
    }
});

// 📄 2. LER PLANILHA
app.get('/api/planilhas/ler', async (req, res) => {
    try {
        const { fileId, driveId } = req.query;
        const client = await getGraphClient();

        const excel = await client
            .api(`/drives/${driveId}/items/${fileId}/workbook/tables/Tabela1/rows`)
            .get();

        const dados = excel.value.map((row, index) => ({
            _id: index,
            data: row.values[0][0],
            tipo: row.values[0][1],
            categoria: row.values[0][2],
            descricao: row.values[0][3],
            valor: row.values[0][4]
        }));

        res.json(dados);

    } catch (err) {
        console.error("❌ Erro ao ler planilha:", err.message);
        res.status(500).json({ erro: err.message });
    }
});

// ✏️ 3. EDITAR LINHA
app.put('/api/planilhas/editar', async (req, res) => {
    try {
        const { fileId, driveId, rowIndex, valores } = req.body;
        const client = await getGraphClient();

        const linhaExcel = rowIndex + 2; // pula cabeçalho

        await client
            .api(`/drives/${driveId}/items/${fileId}/workbook/worksheets('Planilha1')/range(address='A${linhaExcel}:E${linhaExcel}')`)
            .patch({
                values: [valores]
            });

        await registrarLog(req, `Editou linha ${rowIndex} na planilha ${fileId}`);

        res.json({ sucesso: true });

    } catch (err) {
        console.error("❌ Erro ao editar:", err.message);
        res.status(500).json({ erro: err.message });
    }
});

// ➕ 4. INSERIR NOVA LINHA
app.post('/api/planilhas/inserir', async (req, res) => {
    try {
        const { fileId, driveId, data, tipo, categoria, descricao, valor } = req.body;
        const client = await getGraphClient();

        await client
            .api(`/drives/${driveId}/items/${fileId}/workbook/tables/Tabela1/rows/add`)
            .post({
                index: null,
                values: [[data, tipo, categoria, descricao, valor]]
            });

        await registrarLog(req, `Inseriu novo lançamento na planilha ${fileId}`);

        res.json({ sucesso: true });

    } catch (err) {
        console.error("❌ Erro ao inserir:", err.message);
        res.status(500).json({ erro: err.message });
    }
});

app.delete('/api/financeiro/:id', async (req, res) => {
    res.status(400).json({ erro: "Para excluir um lançamento, abra o Excel no seu SharePoint e apague a linha manualmente." });
});

// 3. Inicialização do Servidor
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`🚀 Servidor L2P rodando na porta ${PORT}`));
