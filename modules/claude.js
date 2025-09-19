// modules/claude.js - ENHANCED CLAUDE API MODULE WITH .ENV SUPPORT
const https = require('https');
const fs = require('fs');
const path = require('path');

let CLAUDE_API_KEY = null;
let API_CONFIG = {
    baseURL: 'https://api.anthropic.com/v1/messages',
    model: 'claude-3-haiku-20240307', // Fast and cost-effective model
    maxTokens: 1000
};

let envLoaded = false;

// Load environment variables from .env file
function loadEnvironmentVariables() {
    if (envLoaded) return;
    
    try {
        const envPath = path.join(process.cwd(), '.env');
        
        if (fs.existsSync(envPath)) {
            console.log('📁 Found .env file, loading configuration...');
            
            const envContent = fs.readFileSync(envPath, 'utf8');
            const envLines = envContent.split('\n');
            
            envLines.forEach(line => {
                const trimmedLine = line.trim();
                if (trimmedLine && !trimmedLine.startsWith('#') && trimmedLine.includes('=')) {
                    const [key, ...valueParts] = trimmedLine.split('=');
                    const value = valueParts.join('=').trim();
                    
                    if (value && !process.env[key]) {
                        process.env[key] = value;
                    }
                }
            });
            
            // Load Claude API configuration from environment
            if (process.env.CLAUDE_API_KEY && process.env.CLAUDE_API_KEY.trim()) {
                CLAUDE_API_KEY = process.env.CLAUDE_API_KEY.trim();
                console.log('✅ Claude API key loaded from .env file');
            }
            
            if (process.env.CLAUDE_MODEL) {
                API_CONFIG.model = process.env.CLAUDE_MODEL;
            }
            
            if (process.env.CLAUDE_MAX_TOKENS) {
                API_CONFIG.maxTokens = parseInt(process.env.CLAUDE_MAX_TOKENS) || 1000;
            }
            
            console.log(`🤖 Claude configuration: Model=${API_CONFIG.model}, MaxTokens=${API_CONFIG.maxTokens}`);
            
        } else {
            console.log('📄 No .env file found, will use UI configuration');
            createSampleEnvFile();
        }
        
    } catch (error) {
        console.log('⚠️ Error loading .env file:', error.message);
        console.log('💡 Will use UI configuration instead');
    }
    
    envLoaded = true;
}

// Create sample .env file if it doesn't exist
function createSampleEnvFile() {
    try {
        const envPath = path.join(process.cwd(), '.env.example');
        const sampleEnvContent = `# Windows MCP Chatbot - Environment Configuration
# Copy this file and rename to .env

# Claude API Configuration
CLAUDE_API_KEY=
# Get your API key from: https://console.anthropic.com/

# Optional: Claude API Settings
CLAUDE_MODEL=claude-3-haiku-20240307
CLAUDE_MAX_TOKENS=1000

# Optional: Server Configuration  
PORT=3001

# Instructions:
# 1. Rename this file to .env
# 2. Add your Claude API key
# 3. Restart the server`;

        if (!fs.existsSync(envPath)) {
            fs.writeFileSync(envPath, sampleEnvContent);
            console.log('📝 Created .env.example file for reference');
        }
    } catch (error) {
        console.log('⚠️ Could not create .env.example:', error.message);
    }
}

// Initialize module (load environment on first import)
function initialize() {
    loadEnvironmentVariables();
    
    // Auto-test API if key is loaded from .env
    if (CLAUDE_API_KEY) {
        testAPIConnection()
            .then(status => {
                console.log(`🔗 Claude API auto-test: ${status}`);
            })
            .catch(error => {
                console.log(`⚠️ Claude API auto-test failed: ${error.message}`);
            });
    }
}

// Configure Claude API (now supports both .env and UI)
async function configureAPI(apiKey = null) {
    try {
        // If no API key provided, try to use .env
        if (!apiKey) {
            loadEnvironmentVariables();
            if (CLAUDE_API_KEY) {
                apiKey = CLAUDE_API_KEY;
                console.log('🔄 Using API key from .env file');
            } else {
                throw new Error('No API key provided and none found in .env file');
            }
        }
        
        if (!apiKey || !apiKey.startsWith('sk-ant-')) {
            throw new Error('Invalid Claude API key format. Should start with "sk-ant-"');
        }

        CLAUDE_API_KEY = apiKey;
        
        // Test the API key
        const testResult = await testAPIConnection();
        
        return `✅ Claude API configured successfully!
Source: ${apiKey === process.env.CLAUDE_API_KEY ? '.env file' : 'UI configuration'}
Model: ${API_CONFIG.model}
Status: ${testResult}`;
        
    } catch (error) {
        if (apiKey !== process.env.CLAUDE_API_KEY) {
            CLAUDE_API_KEY = null; // Only reset if not from .env
        }
        throw new Error(`Claude API configuration failed: ${error.message}`);
    }
}

// Test API connection
async function testAPIConnection() {
    if (!CLAUDE_API_KEY) {
        return 'Not configured';
    }

    try {
        const response = await callClaudeAPI('Test connection. Respond with "API working" only.', null, 50);
        return response.includes('API working') || response.includes('working') ? 'Connected and working' : 'Connected but response unclear';
    } catch (error) {
        return `Connection failed: ${error.message}`;
    }
}

// Enhanced status check
async function checkStatus() {
    const envKeyAvailable = !!(process.env.CLAUDE_API_KEY && process.env.CLAUDE_API_KEY.trim());
    const currentKeyConfigured = CLAUDE_API_KEY !== null;
    const keySource = CLAUDE_API_KEY === process.env.CLAUDE_API_KEY ? '.env file' : 'UI configuration';
    
    return {
        configured: currentKeyConfigured,
        keySource: currentKeyConfigured ? keySource : 'not configured',
        envFileExists: fs.existsSync(path.join(process.cwd(), '.env')),
        envKeyAvailable: envKeyAvailable,
        model: API_CONFIG.model,
        maxTokens: API_CONFIG.maxTokens,
        connection: await testAPIConnection()
    };
}

// Main Claude API call function (enhanced with better error handling)
async function callClaudeAPI(prompt, systemPrompt = null, maxTokens = null) {
    if (!CLAUDE_API_KEY) {
        throw new Error('Claude API key not configured. Please add to .env file or configure via UI.');
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
                    const response = JSON.parse(data);
                    
                    if (res.statusCode !== 200) {
                        // Enhanced error messages
                        let errorMessage = `API Error ${res.statusCode}`;
                        if (response.error) {
                            if (response.error.type === 'authentication_error') {
                                errorMessage = 'Invalid API key. Please check your Claude API key.';
                            } else if (response.error.type === 'permission_error') {
                                errorMessage = 'API key does not have required permissions.';
                            } else if (response.error.type === 'rate_limit_error') {
                                errorMessage = 'API rate limit exceeded. Please wait and try again.';
                            } else {
                                errorMessage = response.error.message || 'Unknown API error';
                            }
                        }
                        reject(new Error(errorMessage));
                        return;
                    }

                    if (response.content && response.content[0] && response.content[0].text) {
                        resolve(response.content[0].text);
                    } else {
                        reject(new Error('Invalid response format from Claude API'));
                    }
                } catch (error) {
                    reject(new Error(`Failed to parse API response: ${error.message}`));
                }
            });
        });

        req.on('error', (error) => {
            reject(new Error(`Network error: ${error.message}`));
        });

        // Set timeout for API calls
        req.setTimeout(30000, () => {
            req.destroy();
            reject(new Error('API request timeout (30s)'));
        });

        req.write(postData);
        req.end();
    });
}

// Generate smart email content (enhanced)
async function generateEmailContent(request, recipient, subject) {
    const systemPrompt = `You are an expert email writer. Generate professional, well-structured emails based on user requests. 

Key requirements:
- Keep emails concise but complete (200-400 words typically)
- Use professional but friendly tone
- Include proper greeting and closing
- Format clearly with paragraphs
- Be specific and actionable
- Adapt tone based on email type (formal for leave requests, collaborative for meetings)
- No extra commentary, just the email content`;

    const prompt = `Generate a professional email for this request:

Request: ${request}
To: ${recipient || '[Recipient]'}
Subject: ${subject || '[Subject will be determined]'}

Generate ONLY the email body content (greeting through closing). Be specific and actionable.`;

    try {
        const emailContent = await callClaudeAPI(prompt, systemPrompt);
        return {
            success: true,
            content: emailContent.trim(),
            generatedBy: 'Claude API',
            keySource: CLAUDE_API_KEY === process.env.CLAUDE_API_KEY ? '.env' : 'UI'
        };
    } catch (error) {
        return {
            success: false,
            error: error.message,
            fallback: createFallbackEmail(request, recipient, subject)
        };
    }
}

// Generate smart document content (enhanced)
async function generateDocumentContent(request, documentType) {
    const systemPrompt = `You are an expert writer who creates well-structured documents. Generate appropriate content based on the document type and user request.

Key requirements:
- Create proper document structure with clear sections
- Use appropriate tone for document type (formal for letters, academic for essays)
- Include relevant headings and formatting indicators
- Be comprehensive but focused
- Provide complete, usable content
- No extra commentary, just the document content`;

    const prompt = `Generate a ${documentType} based on this request:

Request: ${request}
Document Type: ${documentType}

Create a complete, well-structured ${documentType} with appropriate sections and content. Include formatting indicators like [HEADING], [PARAGRAPH BREAK] where appropriate.`;

    try {
        const documentContent = await callClaudeAPI(prompt, systemPrompt, 1500); // More tokens for documents
        return {
            success: true,
            content: documentContent.trim(),
            generatedBy: 'Claude API',
            keySource: CLAUDE_API_KEY === process.env.CLAUDE_API_KEY ? '.env' : 'UI'
        };
    } catch (error) {
        return {
            success: false,
            error: error.message,
            fallback: createFallbackDocument(request, documentType)
        };
    }
}

// Handle general queries with Claude (enhanced)
async function handleGeneralQuery(query) {
    try {
        const systemPrompt = `You are a helpful Windows automation assistant with access to Claude AI. Provide clear, actionable responses to user queries about Windows automation, productivity, and general assistance. Keep responses helpful but concise (under 300 words typically).`;
        
        const response = await callClaudeAPI(query, systemPrompt);
        
        return {
            message: '🤖 Claude AI Assistant Response',
            details: `${response}

---
Generated by Claude AI (${CLAUDE_API_KEY === process.env.CLAUDE_API_KEY ? 'configured via .env file' : 'configured via UI'})`
        };
        
    } catch (error) {
        return {
            message: '🤖 Claude AI Assistant (Offline)',
            details: `Claude API not available: ${error.message}

💡 TO FIX THIS:
1. Add your API key to .env file, OR
2. Configure via UI using the "⚙️ Configure" button

I can still help with:
✅ Screenshots and system info
✅ Opening applications  
✅ Basic email and document templates

To enable AI features, get your API key from: https://console.anthropic.com/`
        };
    }
}

// Enhanced email parsing with Claude assistance
async function parseEmailWithClaude(message) {
    try {
        const systemPrompt = `Extract email details from user requests. Return ONLY a JSON object with these fields:
- to: email address (or empty string if not found)
- subject: appropriate subject line based on content
- category: leave/meeting/general/followup/complaint/request
- dates: any dates mentioned (in readable format)
- purpose: brief description of email purpose
- urgency: low/medium/high based on content`;

        const prompt = `Extract email details from: "${message}"

Return only valid JSON, no other text.`;
        
        const response = await callClaudeAPI(prompt, systemPrompt, 200);
        const parsed = JSON.parse(response);
        
        return {
            to: parsed.to || '',
            subject: parsed.subject || 'General Inquiry',
            category: parsed.category || 'general',
            dates: parsed.dates || '',
            purpose: parsed.purpose || '',
            urgency: parsed.urgency || 'medium'
        };
        
    } catch (error) {
        console.log('Claude email parsing failed, using fallback:', error.message);
        return parseEmailBasic(message);
    }
}

// Fallback functions (keep existing ones)
function createFallbackEmail(request, recipient, subject) {
    const msg = request.toLowerCase();
    
    if (msg.includes('leave') || msg.includes('vacation')) {
        return `Dear Sir/Madam,

I would like to request leave for the dates mentioned in my request.

I will ensure all my work is completed before my absence and coordinate with my team for any urgent matters.

Please let me know if you need any additional information.

Thank you for your consideration.

Best regards,
[Your Name]`;
    } else if (msg.includes('meeting')) {
        return `Dear Team,

I would like to schedule a meeting to discuss the topics mentioned in my request.

Please let me know your availability so we can find a suitable time for everyone.

Thank you.

Best regards,
[Your Name]`;
    } else {
        return `Dear ${recipient ? recipient.split('@')[0] : 'Sir/Madam'},

I hope this email finds you well.

[This email was generated automatically. Please review and customize as needed.]

Thank you for your time.

Best regards,
[Your Name]`;
    }
}

function createFallbackDocument(request, documentType) {
    const msg = request.toLowerCase();
    
    if (documentType.includes('letter')) {
        return `Date: ${new Date().toLocaleDateString()}

Dear [Recipient],

I am writing this letter regarding [your request].

[Please customize this content based on your specific needs.]

Thank you for your consideration.

Sincerely,
[Your Name]`;
    } else if (documentType.includes('essay')) {
        return `[Essay Title]

Introduction:
[State your thesis and main points]

Body:
[Develop your arguments with supporting evidence]

Conclusion:
[Summarize your key points]

[This document was auto-generated. Please customize with your specific content.]`;
    } else {
        return `[Document Title]

Created: ${new Date().toLocaleDateString()}

[This document was generated automatically. Please customize with your specific content.]

Content:
[Add your main content here]

Summary:
[Add conclusions or next steps]`;
    }
}

function parseEmailBasic(message) {
    const emailRegex = /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b/g;
    const emails = message.match(emailRegex);
    
    return {
        to: emails ? emails[0] : '',
        subject: message.includes('leave') ? 'Leave Application' : 'General Inquiry',
        category: message.includes('leave') ? 'leave' : 'general',
        dates: '',
        purpose: 'User request',
        urgency: 'medium'
    };
}

// Initialize on module load
initialize();

module.exports = {
    configureAPI,
    checkStatus,
    testAPIConnection,
    callClaudeAPI,
    generateEmailContent,
    generateDocumentContent,
    handleGeneralQuery,
    parseEmailWithClaude,
    loadEnvironmentVariables,
    initialize
};