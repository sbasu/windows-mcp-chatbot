// server.js - BACK TO WORKING VERSION + SMALL FIXES
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
        
        const response = await commands.processUserMessage(message);
        
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

const commands = {
    async processUserMessage(message) {
        const msg = message.toLowerCase();
        
        try {
            // List screenshots
            if (msg.includes('show screenshots') || msg.includes('recent screenshots') || msg.includes('list screenshots')) {
                console.log('📁 Listing screenshots...');
                const result = await this.listRecentScreenshots();
                return {
                    message: '📁 Recent screenshots:',
                    details: result
                };
            }

            // Screenshot - SIMPLE VERSION THAT WORKED
            if (msg.includes('screenshot') || msg.includes('capture')) {
                console.log('📸 Taking screenshot...');
                const result = await this.takeSmartScreenshot();
                return {
                    message: '📸 Screenshot taken!',
                    details: result.success ? result.message : `Error: ${result.error}`
                };
            }

            // App launching - SIMPLE VERSION
            if (msg.includes('open') || msg.includes('launch') || msg.includes('start')) {
                const appName = this.extractAppName(message);
                console.log(`🚀 Launching: ${appName}`);
                const result = await this.launchSimpleApp(appName);
                return {
                    message: `🚀 ${appName} launched!`,
                    details: result.success ? 'Application should be running' : result.error
                };
            }

            // System info
            if (msg.includes('system info') || msg.includes('system')) {
                console.log('💻 Getting system info...');
                const result = await this.getSimpleSystemInfo();
                return {
                    message: '💻 System information:',
                    details: result
                };
            }

            // Default
            return {
                message: '🤖 Simple System Ready!',
                details: 'Commands: screenshot, open calculator, system info, list screenshots'
            };

        } catch (error) {
            console.error('Error in processUserMessage:', error);
            return {
                message: '❌ Command failed',
                details: error.message
            };
        }
    },

    takeSmartScreenshot(type = 'desktop', windowApp = '') {
        return new Promise((resolve) => {
            const timestamp = new Date().toISOString()
                .replace(/T/, '_')
                .replace(/\..+/, '')
                .replace(/:/g, '-');

            const filename = `Screenshot_${type === 'desktop' ? 'Desktop' : type === 'active' ? 'ActiveWindow' : windowApp}_${timestamp}.png`;
            const filepath = path.join(DESKTOP_PATH, filename);

            let scriptContent = '';

            if (type === 'desktop') {
                // Use the SAME working desktop script
                scriptContent = `
Add-Type -AssemblyName System.Windows.Forms,System.Drawing
$screen = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
$bitmap = New-Object System.Drawing.Bitmap($screen.Width, $screen.Height)
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$graphics.CopyFromScreen(0, 0, 0, 0, $bitmap.Size)
$bitmap.Save("${filepath.replace(/\\/g, '\\\\')}")
$graphics.Dispose()
$bitmap.Dispose()
Write-Host "Desktop screenshot saved: ${filepath}"
`;
            } else if (type === 'active') {
                // Active window screenshot
                scriptContent = `
Add-Type -AssemblyName System.Windows.Forms,System.Drawing
Add-Type @'
using System;
using System.Runtime.InteropServices;
public class Win32 {
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    public struct RECT { public int Left, Top, Right, Bottom; }
}
'@

try {
    $hwnd = [Win32]::GetForegroundWindow()
    $rect = New-Object Win32+RECT
    [Win32]::GetWindowRect($hwnd, [ref]$rect) | Out-Null
    $width = $rect.Right - $rect.Left
    $height = $rect.Bottom - $rect.Top
    
    if ($width -gt 0 -and $height -gt 0) {
        $bitmap = New-Object System.Drawing.Bitmap($width, $height)
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        $graphics.CopyFromScreen($rect.Left, $rect.Top, 0, 0, [System.Drawing.Size]::new($width, $height))
        $bitmap.Save("${filepath.replace(/\\/g, '\\\\')}")
        $graphics.Dispose()
        $bitmap.Dispose()
        Write-Host "Active window screenshot saved: ${filepath}"
    } else {
        Write-Error "Could not capture active window - invalid dimensions"
    }
} catch {
    Write-Error "Error capturing active window: $($_.Exception.Message)"
}
`;
            } else if (type === 'window' && windowApp) {
                // Specific application window
                scriptContent = `
Add-Type -AssemblyName System.Windows.Forms,System.Drawing
Add-Type @'
using System;
using System.Runtime.InteropServices;
public class Win32 {
    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    public struct RECT { public int Left, Top, Right, Bottom; }
}
'@

try {
    $processes = Get-Process -Name "${windowApp}" -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 }
    
    if ($processes) {
        $process = $processes[0]
        Write-Host "Found ${windowApp} process with window"
        
        # Bring window to front
        [Win32]::ShowWindow($process.MainWindowHandle, 9) | Out-Null
        [Win32]::SetForegroundWindow($process.MainWindowHandle) | Out-Null
        Start-Sleep -Milliseconds 800
        
        $rect = New-Object Win32+RECT
        [Win32]::GetWindowRect($process.MainWindowHandle, [ref]$rect) | Out-Null
        
        $width = $rect.Right - $rect.Left
        $height = $rect.Bottom - $rect.Top
        
        if ($width -gt 50 -and $height -gt 50) {
            $bitmap = New-Object System.Drawing.Bitmap($width, $height)
            $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
            $graphics.CopyFromScreen($rect.Left, $rect.Top, 0, 0, [System.Drawing.Size]::new($width, $height))
            $bitmap.Save("${filepath.replace(/\\/g, '\\\\')}")
            $graphics.Dispose()
            $bitmap.Dispose()
            Write-Host "${windowApp} window screenshot saved: ${filepath}"
        } else {
            Write-Error "Window dimensions too small: $width x $height"
        }
    } else {
        Write-Error "${windowApp} is not running or has no visible window. Please open ${windowApp} first."
    }
} catch {
    Write-Error "Error capturing ${windowApp} window: $($_.Exception.Message)"
}
`;
            }
            
            // Write script to temp file
            const tempScript = path.join(os.tmpdir(), `screenshot_${Date.now()}.ps1`);
            require('fs').writeFileSync(tempScript, scriptContent);

            console.log(`🔧 Taking ${type} screenshot...`);
            if (type === 'window') console.log(`🎯 Target app: ${windowApp}`);

            // Execute PowerShell script
            exec(`powershell -ExecutionPolicy Bypass -File "${tempScript}"`, (error, stdout, stderr) => {
                // Clean up temp script
                try {
                    require('fs').unlinkSync(tempScript);
                } catch (e) {
                    console.log('Could not delete temp script:', e.message);
                }
                
                if (error || stderr) {
                    console.error('❌ Screenshot error:', error?.message || stderr);
                    resolve({ 
                        success: false, 
                        error: stderr || error?.message || 'Unknown error',
                        message: `Failed to take ${type} screenshot`
                    });
                } else {
                    console.log('✅ Screenshot success!');
                    console.log('📤 Output:', stdout);

                    // Check if file was actually created
                    const fs = require('fs');
                    if (fs.existsSync(filepath)) {
                        const stats = fs.statSync(filepath);
                        resolve({
                            success: true,
                            message: `✅ Saved: ${filename}\n📁 Location: Desktop\n📊 Size: ${Math.round(stats.size/1024)} KB`,
                            filename: filename,
                            filepath: filepath
                        });
                    } else {
                        resolve({
                            success: false,
                            error: 'Screenshot file was not created',
                            message: 'Command completed but file not found'
                        });
                    }
                }
            });
        });
    },

    listRecentScreenshots() {
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
    },

    launchSimpleApp(appName) {
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
    },

    getSimpleSystemInfo() {
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
    },

    extractAppName(message) {
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
};

// Health check
app.get('/api/health', (req, res) => {
    res.json({
        status: 'ok',
        server: 'Back to Working Server',
        desktop_path: DESKTOP_PATH
    });
});

// Start server if not in test environment
if (process.env.NODE_ENV !== 'test') {
    app.listen(PORT, () => {
        console.log(`🚀 BACK TO WORKING Server running on http://localhost:${PORT}`);
        console.log(`📸 Screenshots: Desktop only, no folder opening`);
        console.log(`✅ Back to basics that work!`);
    });
}

process.on('SIGINT', () => {
    console.log('\n🛑 Shutting down...');
    process.exit(0);
});

module.exports = { app, commands };