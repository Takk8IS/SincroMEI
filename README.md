# SincroMEI Google Sheets Add-on

![version](https://img.shields.io/badge/version-1.0.0-blue.svg)
![license](https://img.shields.io/badge/license-CC--BY--4.0-lightgrey.svg)
![node.js](https://img.shields.io/badge/node.js-14.17.0-blue.svg)
![express](https://img.shields.io/badge/express-4.19.2-green.svg)
![dotenv](https://img.shields.io/badge/dotenv-16.4.5-orange.svg)

## Overview

The **SincroMEI Google Sheets Add-on** allows you to synchronise data from the SincroMEI API directly into Google Sheets. This tool is designed to streamline the integration and management of MEI data from the Brazilian Federal Revenue Service.

```mermaid
graph TD
    A["Start Synchronisation"] --> B["Validate CNPJ"]
    B --> C["Send Request to ReceitaWS"]
    C --> D["Receive Response from ReceitaWS"]
    D --> E["Format Received Data"]
    E --> F["Insert Data into Spreadsheet"]
    F --> G["Mark Row as Processed"]
    G --> H["Check Execution Time"]
    H --> I["Wait for Next Batch"]

    I --> J["Synchronization Complete"]
    I --> K["Maximum Execution Time Reached"]
    K --> L["Pause Synchronisation"]
    L --> M["Resume Synchronisation"]

    classDef start fill:#f9f,stroke:#333,stroke-width:2px;
    classDef validate fill:#ffb,stroke:#333,stroke-width:2px;
    classDef send fill:#bbf,stroke:#333,stroke-width:2px;
    classDef receive fill:#bfb,stroke:#333,stroke-width:2px;
    classDef format fill:#bff,stroke:#333,stroke-width:2px;
    classDef insert fill:#fbf,stroke:#333,stroke-width:2px;
    classDef mark fill:#f9f,stroke:#333,stroke-width:2px;
    classDef check fill:#ffb,stroke:#333,stroke-width:2px;
    classDef wait fill:#bbf,stroke:#333,stroke-width:2px;
    classDef complete fill:#9f6,stroke:#333,stroke-width:2px;
    classDef pause fill:#ff6,stroke:#333,stroke-width:2px;
    classDef resume fill:#9f9,stroke:#333,stroke-width:2px;

    class A start;
    class B validate;
    class C send;
    class D receive;
    class E format;
    class F insert;
    class G mark;
    class H check;
    class I wait;
    class J complete;
    class K check;
    class L pause;
    class M resume;

    style A fill:#f9f,stroke:#333,stroke-width:2px;
    style B fill:#ffb,stroke:#333,stroke-width:2px;
    style C fill:#bbf,stroke:#333,stroke-width:2px;
    style D fill:#bfb,stroke:#333,stroke-width:2px;
    style E fill:#bff,stroke:#333,stroke-width:2px;
    style F fill:#fbf,stroke:#333,stroke-width:2px;
    style G fill:#f9f,stroke:#333,stroke-width:2px;
    style H fill:#ffb,stroke:#333,stroke-width:2px;
    style I fill:#bbf,stroke:#333,stroke-width:2px;
    style J fill:#9f6,stroke:#333,stroke-width:2px;
    style K fill:#ffb,stroke:#333,stroke-width:2px;
    style L fill:#ff6,stroke:#333,stroke-width:2px;
    style M fill:#9f9,stroke:#333,stroke-width:2px;
```

## Features

-   **Synchronise MEI Data**: Automatically fetch and populate MEI data in Google Sheets.
-   **Rate Limiting**: Built-in rate limiting to prevent overloading the API.
-   **Data Validation**: Ensures valid CNPJ formats and handles errors gracefully.
-   **Logging**: Comprehensive logging for monitoring and troubleshooting.
-   **Configurable**: Easily set up and customise the spreadsheet.

## Installation

### Prerequisites

-   Node.js 14.17.0 or higher
-   Google Apps Script environment for the Google Sheets add-on
-   A valid API key from ReceitaWS

### Steps

1. **Clone the repository**

    ```bash
    git clone https://github.com/takk8is/sincromei.git
    cd sincromei
    ```

2. **Install dependencies**

    ```bash
    npm install
    ```

3. **Set up environment variables**
   Create a `.env` file in the root directory and add your API key:

    ```
    API_PORT=8001
    API_URL=your_api_url_here
    RECEITAWS_API_KEY=your_api_key_here
    ```

4. **Start the server**

    ```bash
    npm start
    ```

5. **Set up Google Apps Script**
    - Copy the contents of `Code.gs` into a new Google Apps Script project linked to your Google Sheets.
    - Deploy the script as a web app or execute functions directly from the Google Apps Script editor.

## Usage

### Google Sheets Add-on

1. **Open your Google Sheets**
2. **Access the SincroMEI menu**
    - Configure your spreadsheet by selecting `Configurar Planilha`.
    - Start the synchronisation by selecting `Iniciar Sincronização`.
    - Pause the synchronisation by selecting `Pausar Sincronização`.

### API Endpoints

-   **Health Check**

    ```http
    GET /health
    ```

    Check the health status of the API.

-   **Fetch MEI Data**
    ```http
    GET /sincromei/:cnpj
    ```
    Fetch MEI data for a given CNPJ.

## Contributing

We welcome contributions! Please follow these steps:

1. Fork the repository.
2. Create your feature branch (`git checkout -b feature/AmazingFeature`).
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`).
4. Push to the branch (`git push origin feature/AmazingFeature`).
5. Open a Pull Request.

## License

This project is licensed under the Creative Commons Attribution 4.0 International License. See the [LICENSE](LICENSE) file for more details.

## Support

If you have any questions or need support, please open an issue on [GitHub](https://github.com/takk8is/sincromei/issues).

## Donate

Support the project with USDT (TRC-20):

```
TGpiWetnYK2VQpxNGPR27D9vfM6Mei5vNA
```

Feel free to adjust any sections to better fit the specific needs or style preferences of your repository.

## About Takk™ Innovate Studio

Leading the Digital Revolution as the Pioneering 100% Artificial Intelligence Team.

-   Copyright (c) Takk™ Innovate Studio
-   Author: David C Cavalcante
-   Email: say@takk.ag
-   LinkedIn: https://www.linkedin.com/in/hellodav/
-   Medium: https://medium.com/@davcavalcante/
-   Website: https://takk.ag/
-   Twitter: https://twitter.com/takk8is/
-   Medium: https://takk8is.medium.com/
