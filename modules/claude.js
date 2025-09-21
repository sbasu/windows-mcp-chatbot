// modules/claude.js - FIXED .ENV LOADING & API ISSUES
const https = require('https');
const fs = require('fs');
const path = require('path');

let CLAUDE_API_KEY = null;
let API_CONFIG = {
    baseURL: 'https://api.anthropic.com/v1/messages',
    model: 'claude-3-haiku-20240307',
    maxTokens: 1000
};

let envLoaded = false;
let keySource = 'not_configured';

// FIXED: Robust .env file loading
function loadEnvironmentVariables() {
    if (envLoaded) {
        console.log('📁 Environment already loaded');
        return;
    }
    
    try {
        const envPath = path.join(process.cwd(), '.env');
        console.log(`🔍 Looking for .env file at: ${envPath}`);
        
        if (fs.existsSync(envPath)) {
            console.log('✅ .env file found, reading content...');
            
            const envContent = fs.readFileSync(envPath, 'utf8');
            console.log(`📄 .env file size: ${envContent.length} characters`);
            
            const envLines = envContent.split('\n');
            let keyFound = false;
            
            envLines.forEach((line, index) => {
                const trimmedLine = line.trim();
                console.log(`Line ${index + 1}: "${trimmedLine.substring(0, 50)}${trimmedLine.length > 50 ? '...' : ''}"`);
                
                if (trimmedLine && !trimmedLine.startsWith('#') && trimmedLine.includes('=')) {
                    const equalIndex = trimmedLine.indexOf('=');
                    const key = trimmedLine.substring(0, equalIndex).trim();
                    const value = trimmedLine.substring(equalIndex + 1).trim();
                    
                    console.log(`🔑 Found variable: ${key} = ${value.substring(0, 20)}${value.length > 20 ? '...' : ''}`);
                    
                    // Set environment variable if not already set
                    if (!process.env[key]) {
                        process.env[key] = value;
                    }
                    
                    // Check specifically for Claude API key
                    if (key === 'CLAUDE_API_KEY' && value && value.trim() !== '') {
                        CLAUDE_API_KEY = value.trim();
                        keySource = '.env_file';
                        keyFound = true;
                        console.log(`✅ Claude API key loaded from .env: ${CLAUDE_API_KEY.substring(0, 12)}...`);
                    }
                }
            });
            
            if (!keyFound) {
                console.log('⚠️ CLAUDE_API_KEY not found in .env file or is empty');
                console.log('💡 Make sure .env file contains: CLAUDE_API_KEY=sk-ant-your-key-here');
            }
            
            // Load other configuration
            if (process.env.CLAUDE_MODEL) {
                API_CONFIG.model = process.env.CLAUDE_MODEL;
                console.log(`🤖 Claude model set to: ${API_CONFIG.model}`);
            }
            
            if (process.env.CLAUDE_MAX_TOKENS) {
                API_CONFIG.maxTokens = parseInt(process.env.CLAUDE_MAX_TOKENS) || 1000;
                console.log(`📊 Max tokens set to: ${API_CONFIG.maxTokens}`);
            }
            
        } else {
            console.log('📄 No .env file found, creating example...');
            createSampleEnvFile();
        }
        
    } catch (error) {
        console.error('❌ Error loading .env file:', error.message);
        console.log('💡 Will rely on UI configuration instead');
    }
    
    envLoaded = true;
    console.log(`🏁 Environment loading complete. API key source: ${keySource}`);
}

// Create sample .env file with clear instructions
function createSampleEnvFile() {
    try {
        const envExamplePath = path.join(process.cwd(), '.env.example');
        const envRealPath = path.join(process.cwd(), '.env');
        
        const sampleEnvContent = `# Windows MCP Chatbot - Environment Configuration
# IMPORTANT: Rename this file to .env (remove .example)

# Claude API Configuration - GET YOUR KEY FROM: https://console.anthropic.com/
CLAUDE_API_KEY=sk-ant-api03-put-your-actual-key-here

# Optional: Claude API Settings
CLAUDE_MODEL=claude-3-haiku-20240307
CLAUDE_MAX_TOKENS=1000

# Optional: Server Configuration  
PORT=3001
NODE_ENV=development

# INSTRUCTIONS:
# 1. Get your Claude API key from https://console.anthropic.com/
# 2. Replace "sk-ant-api03-put-your-actual-key-here" with your real key
# 3. Save this file as ".env" (remove the .example part)
# 4. Restart the server: node server.js
# 5. The system will automatically load your API key

# SECURITY NOTE:
# Never commit the .env file with real API keys to version control
# Add .env to your .gitignore file`;

        // Create .env.example
        if (!fs.existsSync(envExamplePath)) {
            fs.writeFileSync(envExamplePath, sampleEnvContent);
            console.log('📝 Created .env.example file for reference');
        }
        
        // Create actual .env file with placeholder
        if (!fs.existsSync(envRealPath)) {
            fs.writeFileSync(envRealPath, sampleEnvContent);
            console.log('📝 Created .env file - Please add your API key!');
            console.log('💡 Edit the .env file and add your Claude API key, then restart the server');
        }
        
    } catch (error) {
        console.error('❌ Could not create .env files:', error.message);
    }
}

// FIXED: Better API configuration with validation
async function configureAPI(apiKey = null) {
    try {
        console.log('🔧 Configuring Claude API...');
        
        // If no API key provided, try to use .env
        if (!apiKey || apiKey.trim() === '') {
            console.log('🔍 No API key provided, checking .env file...');
            loadEnvironmentVariables();
            
            if (CLAUDE_API_KEY) {
                apiKey = CLAUDE_API_KEY;
                console.log('✅ Using API key from .env file');
            } else {
                throw new Error('No API key provided and none found in .env file. Please add CLAUDE_API_KEY to your .env file or configure via UI.');
            }
        }
        
        // Validate API key format
        if (!apiKey.startsWith('sk-ant-')) {
            throw new Error(`Invalid Claude API key format. Key should start with "sk-ant-" but received: ${apiKey.substring(0, 10)}...`);
        }
        
        // Set the API key
        CLAUDE_API_KEY = apiKey;
        keySource = apiKey === process.env.CLAUDE_API_KEY ? '.env_file' : 'ui_configuration';
        console.log(`🔑 API key set from: ${keySource}`);
        
        // Test the API key with a simple request
        console.log('🧪 Testing API key...');
        const testResult = await testAPIConnection();
        console.log(`🔗 API test result: ${testResult}`);
        
        return `✅ Claude API configured successfully!

🔑 Key Source: ${keySource}
🤖 Model: ${API_CONFIG.model}
📊 Max Tokens: ${API_CONFIG.maxTokens}
🔗 Connection Status: ${testResult}

${keySource === 'ui_configuration' ? 
'💡 To persist this API key, add it to your .env file:\nCLAUDE_API_KEY=' + apiKey : 
'✅ API key loaded from .env file and working correctly'}`;
        
    } catch (error) {
        // Only reset API key if it wasn't from .env
        if (keySource !== '.env_file') {
            CLAUDE_API_KEY = null;
            keySource = 'not_configured';
        }
        console.error('❌ API configuration failed:', error.message);
        throw new Error(`Claude API configuration failed: ${error.message}`);
    }
}

// FIXED: More robust API connection test
async function testAPIConnection() {
    if (!CLAUDE_API_KEY) {
        return 'Not configured - no API key available';
    }

    try {
        console.log('🔄 Testing Claude API connection...');
        const response = await callClaudeAPI('Reply with exactly: "API test successful"', null, 50);
        console.log(`📨 API response: "${response.substring(0, 50)}..."`);
        
        if (response.toLowerCase().includes('api test successful') || 
            response.toLowerCase().includes('successful') || 
            response.toLowerCase().includes('working')) {
            return 'Connected and working properly';
        } else {
            return `Connected but response unclear: "${response.substring(0, 30)}..."`;
        }
    } catch (error) {
        console.error('❌ API test failed:', error.message);
        return `Connection failed: ${error.message}`;
    }
}

// FIXED: Enhanced status check with detailed information
async function checkStatus() {
    const envKeyAvailable = !!(process.env.CLAUDE_API_KEY && process.env.CLAUDE_API_KEY.trim() && process.env.CLAUDE_API_KEY !== 'sk-ant-api03-put-your-actual-key-here');
    const currentKeyConfigured = CLAUDE_API_KEY !== null;
    const connectionStatus = await testAPIConnection();
    
    return {
        configured: currentKeyConfigured,
        keySource: keySource,
        envFileExists: fs.existsSync(path.join(process.cwd(), '.env')),
        envKeyAvailable: envKeyAvailable,
        envKeyValue: envKeyAvailable ? process.env.CLAUDE_API_KEY.substring(0, 12) + '...' : 'not_set',
        currentKeyValue: currentKeyConfigured ? CLAUDE_API_KEY.substring(0, 12) + '...' : 'not_set',
        model: API_CONFIG.model,
        maxTokens: API_CONFIG.maxTokens,
        connection: connectionStatus,
        debug: {
            envLoaded: envLoaded,
            processEnvKey: process.env.CLAUDE_API_KEY ? process.env.CLAUDE_API_KEY.substring(0, 12) + '...' : 'undefined',
            moduleKey: CLAUDE_API_KEY ? CLAUDE_API_KEY.substring(0, 12) + '...' : 'null'
        }
    };
}

// FIXED: More robust API call with better error handling
async function callClaudeAPI(prompt, systemPrompt = null, maxTokens = null) {
    if (!CLAUDE_API_KEY) {
        throw new Error(`Claude API key not configured. 
        
Please either:
1. Add CLAUDE_API_KEY to your .env file, OR  
2. Configure via UI using the "⚙️ Configure" button

Current status: ${keySource}`);
    }

    const messages = [
        {
            role: 'user',
            content: prompt
        }
    ];

    const payload = {
        model: API_CONFIG.model,
        max_tokens: maxTokens || API_CONFIG.maxTokens,
        messages: messages
    };

    if (systemPrompt) {
        payload.system = systemPrompt;
    }

    return new Promise((resolve, reject) => {
        const postData = JSON.stringify(payload);
        console.log(`🌐 Making API call to Claude (${payload.model}, ${payload.max_tokens} tokens)...`);
        
        const options = {
            hostname: 'api.anthropic.com',
            port: 443,
            path: '/v1/messages',
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Content-Length': Buffer.byteLength(postData),
                'x-api-key': CLAUDE_API_KEY,
                'anthropic-version': '2023-06-01'
            }
        };

        const req = https.request(options, (res) => {
            let data = '';

            res.on('data', (chunk) => {
                data += chunk;
            });

            res.on('end', () => {
                try {
                    console.log(`📨 API response status: ${res.statusCode}`);
                    const response = JSON.parse(data);
                    
                    if (res.statusCode !== 200) {
                        let errorMessage = `API Error ${res.statusCode}`;
                        if (response.error) {
                            switch (response.error.type) {
                                case 'authentication_error':
                                    errorMessage = `Invalid API key. Please check your Claude API key.
                                    
Current key: ${CLAUDE_API_KEY.substring(0, 12)}...
Key source: ${keySource}

To fix:
1. Get valid key from https://console.anthropic.com/
2. Update .env file or reconfigure via UI`;
                                    break;
                                case 'permission_error':
                                    errorMessage = 'API key does not have required permissions. Contact Anthropic support.';
                                    break;
                                case 'rate_limit_error':
                                    errorMessage = 'API rate limit exceeded. Please wait and try again.';
                                    break;
                                default:
                                    errorMessage = response.error.message || 'Unknown API error';
                            }
                        }
                        reject(new Error(errorMessage));
                        return;
                    }

                    if (response.content && response.content[0] && response.content[0].text) {
                        console.log('✅ API call successful');
                        resolve(response.content[0].text);
                    } else {
                        reject(new Error('Invalid response format from Claude API'));
                    }
                } catch (error) {
                    reject(new Error(`Failed to parse API response: ${error.message}\nRaw response: ${data.substring(0, 200)}`));
                }
            });
        });

        req.on('error', (error) => {
            reject(new Error(`Network error: ${error.message}. Check your internet connection.`));
        });

        req.setTimeout(30000, () => {
            req.destroy();
            reject(new Error('API request timeout (30 seconds). Try again or check your connection.'));
        });

        req.write(postData);
        req.end();
    });
}

// Initialize on module load
function initialize() {
    console.log('🚀 Initializing Claude module...');
    loadEnvironmentVariables();
    
    // Auto-test API if key is loaded from .env
    if (CLAUDE_API_KEY && keySource === '.env_file') {
        console.log('🧪 Auto-testing API key from .env...');
        testAPIConnection()
            .then(status => {
                console.log(`✅ Auto-test result: ${status}`);
            })
            .catch(error => {
                console.error(`❌ Auto-test failed: ${error.message}`);
            });
    } else {
        console.log('💡 No API key in .env file. Configure via UI or add to .env file.');
    }
}

// Generate smart email content (keep existing)
async function generateEmailContent(request, recipient, subject) {
    const systemPrompt = `You are an expert email writer. Generate professional, well-structured emails based on user requests.`;
    const prompt = `Generate a professional email for this request: ${request}`;

    try {
        const emailContent = await callClaudeAPI(prompt, systemPrompt);
        return {
            success: true,
            content: emailContent.trim(),
            generatedBy: 'Claude API',
            keySource: keySource
        };
    } catch (error) {
        return {
            success: false,
            error: error.message,
            fallback: `Dear ${recipient || 'Sir/Madam'},

I hope this email finds you well.

[This email was auto-generated. Please customize as needed.]

Best regards,
[Your Name]`
        };
    }
}

// Generate smart document content (keep existing)
async function generateDocumentContent(request, documentType) {
    const systemPrompt = `You are an expert writer. Create well-structured documents based on the request and document type.`;
    const prompt = `Generate a ${documentType} based on this request: ${request}`;

    try {
        const documentContent = await callClaudeAPI(prompt, systemPrompt, 1500);
        return {
            success: true,
            content: documentContent.trim(),
            generatedBy: 'Claude API',
            keySource: keySource
        };
    } catch (error) {
        return {
            success: false,
            error: error.message,
            fallback: `Document created on ${new Date().toLocaleDateString()}

[This document was auto-generated. Please customize as needed.]`
        };
    }
}

// Handle general queries (keep existing)
async function handleGeneralQuery(query) {
    try {
        const systemPrompt = `You are a helpful Windows automation assistant. Provide clear, actionable responses.`;
        const response = await callClaudeAPI(query, systemPrompt);
        
        return {
            message: '🤖 Claude AI Assistant Response',
            details: `${response}

---
Generated by Claude API (${keySource})`
        };
    } catch (error) {
        return {
            message: '🤖 Claude AI Assistant (Offline)',
            details: `Claude API not available: ${error.message}

🔧 TO FIX:
1. Check your .env file contains: CLAUDE_API_KEY=your-key-here
2. Or configure via UI using "⚙️ Configure" button
3. Get API key from: https://console.anthropic.com/

Current status: ${keySource}
Working features: Screenshots, System info, Basic templates`
        };
    }
}

// Initialize immediately when module is loaded
initialize();

module.exports = {
    configureAPI,
    checkStatus,
    testAPIConnection,
    callClaudeAPI,
    generateEmailContent,
    generateDocumentContent,
    handleGeneralQuery,
    loadEnvironmentVariables,
    initialize
};