// modules/screenshot.js - SCREENSHOT MODULE (WORKING)
const { exec } = require('child_process');
const path = require('path');
const os = require('os');

const DESKTOP_PATH = path.join(os.homedir(), 'Desktop');

async function handleScreenshot(message) {
    try {
        const result = await takeDesktopScreenshot();
        
        if (result.success) {
            // Also list recent screenshots
            const recentList = await listRecentScreenshots();
            
            return {
                message: '📸 Screenshot taken successfully!',
                details: `${result.message}\n\n📁 Recent screenshots:\n${recentList}`
            };
        } else {
            return {
                message: '❌ Screenshot failed',
                details: result.error
            };
        }
    } catch (error) {
        return {
            message: '❌ Screenshot error',
            details: error.message
        };
    }
}

function takeDesktopScreenshot() {
    return new Promise((resolve) => {
        const timestamp = new Date().toISOString()
            .replace(/T/, '_')
            .replace(/\..+/, '')
            .replace(/:/g, '-');

        const filename = `Screenshot_Desktop_${timestamp}.png`;
        const filepath = path.join(DESKTOP_PATH, filename);

        // Simple, working PowerShell screenshot script
        const scriptContent = `
Add-Type -AssemblyName System.Windows.Forms,System.Drawing
$screen = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
$bitmap = New-Object System.Drawing.Bitmap($screen.Width, $screen.Height)
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$graphics.CopyFromScreen(0, 0, 0, 0, $bitmap.Size)
$bitmap.Save("${filepath.replace(/\\/g, '\\\\')}")
$graphics.Dispose()
$bitmap.Dispose()
Write-Host "Screenshot saved successfully to Desktop"`;

        const tempScript = path.join(os.tmpdir(), `screenshot_${Date.now()}.ps1`);
        require('fs').writeFileSync(tempScript, scriptContent);

        console.log('📸 Taking desktop screenshot...');

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
                    error: stderr || error?.message || 'Unknown screenshot error'
                });
            } else {
                console.log('✅ Screenshot success!');
                
                // Verify file was created
                const fs = require('fs');
                if (fs.existsSync(filepath)) {
                    const stats = fs.statSync(filepath);
                    resolve({
                        success: true,
                        message: `✅ Screenshot saved: ${filename}\n📁 Location: Desktop\n📊 File size: ${Math.round(stats.size/1024)} KB`,
                        filename: filename,
                        filepath: filepath
                    });
                } else {
                    resolve({
                        success: false,
                        error: 'Screenshot command completed but file was not created'
                    });
                }
            }
        });
    });
}

function listRecentScreenshots() {
    return new Promise((resolve) => {
        const command = `powershell -Command "Get-ChildItem '${DESKTOP_PATH}\\Screenshot*.png' | Sort-Object LastWriteTime -Descending | Select-Object -First 5 | ForEach-Object { Write-Host ($_.Name + ' - ' + $_.LastWriteTime.ToString('MM/dd HH:mm') + ' (' + [math]::Round($_.Length/1KB, 1) + ' KB)') }"`;

        exec(command, (error, stdout, stderr) => {
            if (error || !stdout.trim()) {
                resolve('No previous screenshots found on Desktop');
            } else {
                resolve(stdout.trim());
            }
        });
    });
}

module.exports = {
    handleScreenshot,
    takeDesktopScreenshot,
    listRecentScreenshots
};