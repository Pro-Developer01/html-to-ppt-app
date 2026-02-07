import puppeteer from 'puppeteer';
import http from 'http';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const PORT = 3001;

// Simple HTTP server
function startServer() {
    return new Promise((resolve) => {
        const mimeTypes = {
            '.html': 'text/html',
            '.js': 'text/javascript',
            '.mjs': 'text/javascript',
            '.css': 'text/css',
            '.woff': 'font/woff',
            '.woff2': 'font/woff2',
        };

        const server = http.createServer((req, res) => {
            let filePath = '.' + req.url;
            if (filePath === './') filePath = './converter-auto.html';

            const extname = String(path.extname(filePath)).toLowerCase();
            const contentType = mimeTypes[extname] || 'application/octet-stream';

            fs.readFile(filePath, (error, content) => {
                if (error) {
                    res.writeHead(404);
                    res.end('Not found');
                } else {
                    res.writeHead(200, { 
                        'Content-Type': contentType,
                        'Access-Control-Allow-Origin': '*'
                    });
                    res.end(content, 'utf-8');
                }
            });
        });

        server.listen(PORT, () => {
            console.log(`✓ Server started on http://localhost:${PORT}`);
            resolve(server);
        });
    });
}

async function convertWithPuppeteer() {
    console.log('Converting HTML slides to editable PowerPoint using dom-to-pptx...\n');

    const server = await startServer();

    try {
        console.log('✓ Launching browser...');
        const browser = await puppeteer.launch({ 
            headless: true,
            args: ['--no-sandbox', '--disable-setuid-sandbox']
        });
        
        const page = await browser.newPage();
        
        // Set download behavior
        const downloadPath = path.resolve(__dirname);
        const client = await page.createCDPSession();
        await client.send('Page.setDownloadBehavior', {
            behavior: 'allow',
            downloadPath: downloadPath
        });

        console.log('✓ Opening converter page...');
        await page.goto(`http://localhost:${PORT}/converter-auto.html`, {
            waitUntil: 'networkidle0',
            timeout: 60000
        });

        console.log('✓ Running conversion...');
        
        // Wait for conversion to complete
        await page.waitForFunction(
            () => document.getElementById('status').textContent.includes('Success') || 
                  document.getElementById('status').textContent.includes('Error'),
            { timeout: 120000 }
        );

        const status = await page.evaluate(() => document.getElementById('status').textContent);
        console.log('\nConversion Status:');
        console.log(status);

        // Wait a bit for downloads to complete
        await new Promise(resolve => setTimeout(resolve, 3000));

        await browser.close();
        console.log('\n✅ Browser closed');

    } catch (error) {
        console.error('Error:', error);
    } finally {
        server.close();
        console.log('✓ Server stopped');
    }
}

convertWithPuppeteer().catch(console.error);
