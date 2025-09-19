// server.js - ENHANCED WITH CLAUDE API INTEGRATION
const express = require('express');
const cors = require('cors');
const path = require('path');
const os = require('os');

// Import separate modules
const screenshotModule = require('./modules/screenshot');
const appLauncherModule = require('./modules/appLauncher');
const systemInfoModule = require('./modules/systemInfo');
const emailModule = require('./modules/email');
const wordModule = require('./modules/word');
const claudeModule = require('./modules/claude'); // NEW: Claude API module

const app = express();
const PORT = 3001;

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.static('public'));

// Routes
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Main chat endpoint
app.post('/api/chat', async (req, res) => {
    try {
        const { message } = req.body;
        console.log(`📨 Message: ${message}`);
        
        const response = await processMessage(message);
        
        res.json({
            success: true,
            response: response.message,
            details: response.details,
            timestamp: new Date().toISOString()
        });
        
    } catch (error) {
        console.error('❌ Error:', error);
        res.status(500).json({
            success: false,
            error: error.message
        });
    }
});

// Enhanced message processor with Claude API integration
async function processMessage(message) {
    const msg = message.toLowerCase();
    
    try {
        // CLAUDE API POWERED FEATURES 🤖
        
        // Smart email composition with Claude API
        if (msg.includes('email') || msg.includes('compose')) {
            console.log('📧 Processing smart email with Claude API...');
            return await emailModule.handleSmartEmail(message, claudeModule);
        }

        // Smart document creation with Claude API
        if (msg.includes('word') || (msg.includes('write') && (msg.includes('document') || msg.includes('letter') || msg.includes('essay')))) {
            console.log('📝 Processing smart document with Claude API...');
            return await wordModule.handleSmartWord(message, claudeModule);
        }

        // Claude-powered general assistance
        if (msg.includes('help') || msg.includes('assistant') || msg.includes('claude')) {
            console.log('🤖 Using Claude for general assistance...');
            return await claudeModule.handleGeneralQuery(message);
        }

        // WORKING FEATURES (unchanged) ✅
        
        // Screenshot (WORKING)
        if (msg.includes('screenshot') || msg.includes('capture')) {
            console.log('📸 Taking screenshot...');
            return await screenshotModule.handleScreenshot(message);
        }

        // App launching (WORKING for basic apps)
        if (msg.includes('open') || msg.includes('launch') || msg.includes('start')) {
            const appName = extractAppName(message);
            console.log(`🚀 Launching: ${appName}`);
            return await appLauncherModule.handleAppLaunch(appName);
        }

        // System info (WORKING)
        if (msg.includes('system info') || msg.includes('system')) {
            console.log('💻 Getting system info...');
            return await systemInfoModule.handleSystemInfo();
        }

        // Default response
        return {
            message: '🤖 Claude-Enhanced Windows Assistant Ready',
            details: 'Features: Smart emails, Smart documents, Screenshots, App launching\nNow with Claude API for intelligent content generation!'
        };

    } catch (error) {
        console.error('Error in processMessage:', error);
        return {
            message: '❌ Command failed',
            details: error.message
        };
    }
}

// Helper function
function extractAppName(message) {
    const msg = message.toLowerCase();
    const workingApps = ['calculator', 'notepad', 'paint', 'cmd', 'powershell'];
    const experimentalApps = ['outlook', 'word', 'excel', 'chrome', 'edge'];
    
    // Check working apps first
    for (const app of workingApps) {
        if (msg.includes(app)) return app;
    }
    
    // Then check experimental apps
    for (const app of experimentalApps) {
        if (msg.includes(app)) return app;
    }
    
    return 'calculator'; // safe default
}

// New endpoint for Claude API configuration
app.post('/api/configure-claude', async (req, res) => {
    try {
        const { apiKey } = req.body;
        
        if (!apiKey) {
            return res.status(400).json({
                success: false,
                error: 'API key is required'
            });
        }

        const result = await claudeModule.configureAPI(apiKey);
        
        res.json({
            success: true,
            message: 'Claude API configured successfully',
            details: result
        });
        
    } catch (error) {
        res.status(500).json({
            success: false,
            error: error.message
        });
    }
});

// Endpoint to check Claude API status
app.get('/api/claude-status', async (req, res) => {
    try {
        const status = await claudeModule.checkStatus();
        res.json({
            success: true,
            status: status
        });
    } catch (error) {
        res.status(500).json({
            success: false,
            error: error.message
        });
    }
});

// Health check
app.get('/api/health', (req, res) => {
    res.json({
        status: 'ok',
        server: 'Claude-Enhanced Windows Assistant',
        working_features: ['Screenshot', 'Basic App Launch', 'System Info'],
        claude_features: ['Smart Email Composition', 'Smart Document Creation', 'General AI Assistant'],
        desktop_path: path.join(os.homedir(), 'Desktop')
    });
});

// Start server
if (process.env.NODE_ENV !== 'test') {
    app.listen(PORT, () => {
        console.log(`🚀 CLAUDE-ENHANCED Server running on http://localhost:${PORT}`);
        console.log(`✅ Working: Screenshots, Calculator, Notepad, System Info`);
        console.log(`🤖 Claude API: Smart emails, documents, and general assistance`);
        console.log(`🔧 Configure Claude API key via frontend or POST /api/configure-claude`);
    });
}

process.on('SIGINT', () => {
    console.log('\n🛑 Shutting down...');
    process.exit(0);
});

module.exports = app;