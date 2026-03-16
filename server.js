// 1. Importações e Configurações Iniciais
require('dotenv').config(); // Carrega o .env se estiver rodando localmente
const express = require('express');
const mongoose = require('mongoose');
const cors = require('cors');

const app = express();
app.use(cors({ origin: '*' })); 
app.use(express.json());

// 2. Conexão Segura com MongoDB
// O process.env.MONGO_URI buscará a chave no Render ou no seu arquivo .env local
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

// --- MODELOS ---
const Apropriacao = mongoose.model('Apropriacao', new mongoose.Schema({ data: String, dados_dia: Object }));
const Funcionario = mongoose.model('Funcionario', new mongoose.Schema({ mat: String, nome: String }));
const Projeto = mongoose.model('Projeto', new mongoose.Schema({ codigo: String }));

// --- FUNÇÃO DE SEMENTE (DATABASE SEED) ---
async function semearBanco() {
    const fCount = await Funcionario.countDocuments();
    if (fCount === 0) {
        const iniciais = [
            { mat: "79", nome: "Amarildo Fernandes Rosa" }, { mat: "56", nome: "André Luis Teixeira" },
            { mat: "88", nome: "Anderson dos Santos Silva" }, { mat: "27", nome: "Cicero Fernandes De Morais" },
            { mat: "91", nome: "Diogo Bassi Rosa" }, { mat: "73", nome: "Douglas Silva Neves" },
            { mat: "38", nome: "Edeilson Bezerra Rocha" }, { mat: "57", nome: "Edmundo Vilas Boas Filho" }
        ];
        await Funcionario.insertMany(iniciais);
        await Projeto.insertMany([{codigo:'230304'}, {codigo:'242236'}, {codigo:'C522006'}, {codigo:'ATESTADO'}]);
        console.log('✅ Dados iniciais semeados no banco de dados!');
    }
}

// --- ROTAS DA API ---

// Apropriação de Horas
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

// Gestão de Equipe
app.get('/api/equipe', async (req, res) => res.json(await Funcionario.find().sort({ nome: 1 })));

app.post('/api/equipe', async (req, res) => { 
    try {
        await new Funcionario(req.body).save(); 
        res.json({sucesso:true}); 
    } catch (err) { res.status(500).json({ erro: err.message }); }
});

app.delete('/api/equipe/:mat', async (req, res) => { 
    await Funcionario.deleteOne({ mat: req.params.mat }); 
    res.json({sucesso:true}); 
});

// Gestão de Projetos
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

// Rota de Login Segura
app.post('/api/login', (req, res) => {
    const { usuario, senha } = req.body;

    // As senhas ficam seguras aqui no backend!
    if (usuario === 'gerencia' && senha === 'L2pgerencia2026!') {
        res.json({ tipo: 'admin', nome: 'Gerência' });
    } 
    else if (usuario === 'analista' && senha === 'L2panalista2026') {
        res.json({ tipo: 'user', nome: 'Analista L2P' });
    } 
    else {
        // Retorna erro 401 (Não autorizado) sem expor a senha correta
        res.status(401).json({ erro: 'Usuário ou senha incorretos' });
    }
});

// 3. Inicialização do Servidor
// Importante: O Render exige que a porta seja dinâmica (process.env.PORT)
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`🚀 Servidor L2P rodando na porta ${PORT}`));
