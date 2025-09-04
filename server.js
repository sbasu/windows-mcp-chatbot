// server.js - SUPER SIMPLE VERSION WITH WORKING SCREENSHOTS
const express = require('express');
const cors = require('cors');
const path = require('path');
const { exec } = require('child_process');
const os = require('os');

const app = express();
const PORT = 3001;

// Get user's desktop path
const DESKTOP_PATH = path.join(os.homedir(), 'Desktop');

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
        
        const response = await processUserMessage(message);
        
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

// Process messages
async function processUserMessage(message) {
    const msg = message.toLowerCase();
    
    try {
        // Screenshot
        if (msg.includes('screenshot') || msg.includes('capture')) {
            console.log('📸 Taking screenshot...');
            const result = await takeSimpleScreenshot();
            return {
                message: '📸 Screenshot taken!',
                details: result.success ? result.message : `Error: ${result.error}`
            };
        }
        
        // List screenshots
        if (msg.includes('show screenshots') || msg.includes('recent screenshots') || msg.includes('list screenshots')) {
            console.log('📁 Listing screenshots...');
            const result = await listRecentScreenshots();
            return {
                message: '📁 Recent screenshots:',
                details: result
            };
        }
        
        // App launching
        if (msg.includes('open') || msg.includes('launch') || msg.includes('start')) {
            const appName = extractAppName(message);
            console.log(`🚀 Launching: ${appName}`);
            const result = await launchSimpleApp(appName);
            return {
                message: `🚀 ${appName} launched!`,
                details: result.success ? 'Application should be running' : result.error
            };
        }
        
        // System info
        if (msg.includes('system info') || msg.includes('system')) {
            console.log('💻 Getting system info...');
            const result = await getSimpleSystemInfo();
            return {
                message: '💻 System information:',
                details: result
            };
        }
        
        // Default
        return {
            message: '🤖 Super Simple System Ready!',
            details: 'Commands: screenshot, open calculator, system info, list screenshots'
        };
        
    } catch (error) {
        console.error('Error in processUserMessage:', error);
        return {
            message: '❌ Command failed',
            details: error.message
        };
    }
}

// SUPER SIMPLE screenshot function that WILL work
function takeSimpleScreenshot() {
    return new Promise((resolve) => {
        const timestamp = new Date().toISOString()
            .replace(/T/, '_')
            .replace(/\..+/, '')
            .replace(/:/g, '-');
        
        const filename = `Screenshot_${timestamp}.png`;
        const filepath = path.join(DESKTOP_PATH, filename);
        
        // Create a temporary PowerShell script file to avoid command line escaping issues
        const scriptContent = `
Add-Type -AssemblyName System.Windows.Forms,System.Drawing
$screen = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
$bitmap = New-Object System.Drawing.Bitmap($screen.Width, $screen.Height)
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$graphics.CopyFromScreen(0, 0, 0, 0, $bitmap.Size)
$bitmap.Save("${filepath.replace(/\\/g, '\\\\')}")
$graphics.Dispose()
$bitmap.Dispose()
Write-Host "Screenshot saved: ${filepath}"
Start-Process explorer.exe -ArgumentList "${DESKTOP_PATH.replace(/\\/g, '\\\\')}"
`;
        
        // Write script to temp file
        const tempScript = path.join(os.tmpdir(), `screenshot_${Date.now()}.ps1`);
        require('fs').writeFileSync(tempScript, scriptContent);
        
        console.log(`🔧 Taking screenshot using temp script: ${tempScript}`);
        console.log(`📁 Will save to: ${filepath}`);
        
        // Execute PowerShell script
        exec(`powershell -ExecutionPolicy Bypass -File "${tempScript}"`, (error, stdout, stderr) => {
            // Clean up temp script
            try {
                require('fs').unlinkSync(tempScript);
            } catch (e) {
                console.log('Could not delete temp script:', e.message);
            }
            
            if (error) {
                console.error('❌ Screenshot error:', error.message);
                console.error('❌ stderr:', stderr);
                resolve({ 
                    success: false, 
                    error: error.message,
                    message: 'Failed to take screenshot'
                });
            } else {
                console.log('✅ Screenshot success!');
                console.log('📤 stdout:', stdout);
                
                // Check if file was actually created
                const fs = require('fs');
                if (fs.existsSync(filepath)) {
                    const stats = fs.statSync(filepath);
                    resolve({ 
                        success: true, 
                        message: `✅ Saved: ${filename}\n📁 Location: Desktop\n📊 Size: ${Math.round(stats.size/1024)} KB\n🗂️ Desktop opened`,
                        filename: filename,
                        filepath: filepath
                    });
                } else {
                    resolve({ 
                        success: false, 
                        error: 'File was not created',
                        message: 'Screenshot command ran but file not found'
                    });
                }
            }
        });
    });
}

// List recent screenshots
function listRecentScreenshots() {
    return new Promise((resolve) => {
        const command = `powershell -Command "Get-ChildItem '${DESKTOP_PATH}\\Screenshot*.png' | Sort-Object LastWriteTime -Descending | Select-Object -First 5 | ForEach-Object { Write-Host ($_.Name + ' - ' + $_.LastWriteTime.ToString('MM/dd HH:mm') + ' (' + [math]::Round($_.Length/1KB, 1) + ' KB)') }"`;
        
        exec(command, (error, stdout, stderr) => {
            if (error || !stdout.trim()) {
                resolve('No screenshots found on Desktop');
            } else {
                resolve(stdout.trim());
            }
        });
    });
}

// Simple app launcher
function launchSimpleApp(appName) {
    return new Promise((resolve) => {
        const apps = {
            'calculator': 'calc',
            'notepad': 'notepad',
            'paint': 'mspaint',
            'task manager': 'taskmgr',
            'explorer': 'explorer',
            'cmd': 'cmd',
            'powershell': 'powershell',
            'word': 'winword',
            'excel': 'excel',
            'edge': 'msedge',
            'chrome': 'chrome',
            'firefox': 'firefox'
        };
        
        const executable = apps[appName.toLowerCase()] || appName;
        const command = `start "" "${executable}"`;
        
        exec(command, (error, stdout, stderr) => {
            if (error) {
                resolve({ success: false, error: error.message });
            } else {
                resolve({ success: true });
            }
        });
    });
}

// Simple system info
function getSimpleSystemInfo() {
    return new Promise((resolve) => {
        const command = 'powershell -Command "Get-ComputerInfo | Select-Object WindowsProductName, @{Name=\\"Memory_GB\\";Expression={[math]::Round($_.TotalPhysicalMemory/1GB,1)}} | Format-List"';
        
        exec(command, (error, stdout, stderr) => {
            if (error) {
                resolve(`Error getting system info: ${error.message}`);
            } else {
                resolve(stdout.trim() || 'System info not available');
            }
        });
    });
}

// Extract app name from message
function extractAppName(message) {
    const msg = message.toLowerCase();
    if (msg.includes('calculator')) return 'calculator';
    if (msg.includes('notepad')) return 'notepad';
    if (msg.includes('paint')) return 'paint';
    if (msg.includes('task manager')) return 'task manager';
    if (msg.includes('explorer')) return 'explorer';
    if (msg.includes('cmd')) return 'cmd';
    if (msg.includes('powershell')) return 'powershell';
    if (msg.includes('word')) return 'word';
    if (msg.includes('excel')) return 'excel';
    if (msg.includes('edge')) return 'edge';
    if (msg.includes('chrome')) return 'chrome';
    if (msg.includes('firefox')) return 'firefox';
    return 'calculator';
}

// Health check
app.get('/api/health', (req, res) => {
    res.json({
        status: 'ok',
        server: 'Super Simple Server with WORKING Screenshots',
        desktop_path: DESKTOP_PATH
    });
});

// Start server
app.listen(PORT, () => {
    console.log(`🚀 SUPER SIMPLE Server running on http://localhost:${PORT}`);
    console.log(`📸 Screenshots will be saved to: ${DESKTOP_PATH}`);
    console.log(`✅ Using temporary PowerShell scripts for reliability`);
    console.log(`🔧 This version WILL work!`);
});

process.on('SIGINT', () => {
    console.log('\n🛑 Shutting down...');
    process.exit(0);
});