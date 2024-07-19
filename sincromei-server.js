// ```sincromei-server.js

/**
 * Copyright (c) Takk™ Innovate Studio
 * SincroMEI Google Sheets Add-on
 * Description: This Google Sheets add-on allows you to synchronise data from the SincroMEI API.
 * Author: David C Cavalcante
 * Version: 1.0.0
 * License: CC-BY-4.0
 * Repository: https://github.com/takk8is/sincromei.git
 * Donate: $USDT (TRC-20) TGpiWetnYK2VQpxNGPR27D9vfM6Mei5vNA
 */

// Import the required modules
import "dotenv/config";
import express from "express";
import helmet from "helmet";
import cors from "cors";
import rateLimit from "express-rate-limit";
import axios from "axios";
import winston from "winston";
import morgan from "morgan";

const app = express();
const port = process.env.API_PORT || 8001;
const apiKey = process.env.RECEITAWS_API_KEY;

// Logger configuration
const logger = winston.createLogger({
    level: "info",
    format: winston.format.json(),
    defaultMeta: { service: "sincromei-service" },
    transports: [
        new winston.transports.File({ filename: "error.log", level: "error" }),
        new winston.transports.File({ filename: "combined.log" }),
    ],
});

if (process.env.NODE_ENV !== "production") {
    logger.add(
        new winston.transports.Console({
            format: winston.format.simple(),
        }),
    );
}

// Middleware
app.use(helmet());
app.use(cors());
app.use(express.json());
app.use(morgan("combined"));

// Rate limiting
const limiter = rateLimit({
    // 15 minutes
    windowMs: 15 * 60 * 1000,
    // limit each IP to 100 requests per windowMs
    max: 100,
});
app.use(limiter);

// CNPJ validation middleware
const validateCNPJ = (req, res, next) => {
    const cnpj = req.params.cnpj;
    if (!/^\d{14}$/.test(cnpj)) {
        return res.status(400).json({ error: "Invalid CNPJ format" });
    }
    next();
};

// Helper function to format years data
const formatYears = (situacao) => {
    const currentYear = new Date().getFullYear();
    const years = {};
    for (let year = 2019; year <= currentYear; year++) {
        years[year] = year === currentYear ? situacao : "Não Optante";
    }
    return Object.entries(years).map(([year, status]) => `${year}:${status}`);
};

// Main route
app.get("/sincromei/:cnpj", validateCNPJ, async (req, res) => {
    const cnpj = req.params.cnpj;
    const url = `https://www.receitaws.com.br/v1/cnpj/${cnpj}`;

    try {
        const response = await axios.get(url, {
            headers: {
                Authorization: `Bearer ${apiKey}`,
            },
        });

        const result = response.data;
        const data = {
            status: result.status,
            ultima_atualizacao: result.ultima_atualizacao,
            cnpj: result.cnpj,
            tipo: result.tipo,
            porte: result.porte,
            nome: result.nome,
            fantasia: result.fantasia,
            abertura: result.abertura,
            atividade_principal: result.atividade_principal,
            atividades_secundarias: result.atividades_secundarias,
            natureza_juridica: result.natureza_juridica,
            logradouro: result.logradouro,
            numero: result.numero,
            complemento: result.complemento,
            cep: result.cep,
            bairro: result.bairro,
            municipio: result.municipio,
            uf: result.uf,
            email: result.email,
            telefone: result.telefone,
            efr: result.efr,
            situacao: result.situacao,
            data_situacao: result.data_situacao,
            motivo_situacao: result.motivo_situacao,
            situacao_especial: result.situacao_especial,
            data_situacao_especial: result.data_situacao_especial,
            capital_social: result.capital_social,
            qsa: result.qsa,
            billing: result.billing,
            anos: formatYears(result.situacao),
        };

        res.json(data);
    } catch (error) {
        logger.error("Error fetching data from ReceitaWS", {
            error: error.message,
            cnpj,
        });
        res.status(500).json({ error: "Error fetching data" });
    }
});

// Health check route
app.get("/health", (req, res) => {
    res.status(200).json({ status: "OK" });
});

// Error handling middleware
app.use((err, req, res, next) => {
    logger.error("Unhandled error", { error: err.message });
    res.status(500).json({ error: "Internal server error" });
});

// Start the server
const server = app.listen(port, () => {
    logger.info(`Server running on port ${port}`);
});

// Graceful shutdown
process.on("SIGTERM", () => {
    logger.info("SIGTERM signal received: closing HTTP server");
    server.close(() => {
        logger.info("HTTP server closed");
        process.exit(0);
    });
});
