// modules/appLauncher.js - APP LAUNCHER MODULE (WORKING)
const { exec } = require('child_process');

async function handleAppLaunch(appName) {
    try {
        const result = await launchApplication(appName);
        
        return {
            message: `🚀 ${appName} launch result`,
            details: result
        };
    } catch (error) {
        return {
            message: `❌ Failed to launch ${appName}`,
            details: error.message
        };
    }
}

function launchApplication(appName) {
    return new Promise((resolve) => {
        // Define apps with their status (working vs experimental)
        const apps = {
            // WORKING APPS ✅ (basic Windows apps)
            'calculator': { 
                cmd: 'calc', 
                name: 'Calculator',
                status: 'working',
                description: 'Windows Calculator'
            },
            'notepad': { 
                cmd: 'notepad', 
                name: 'Notepad',
                status: 'working',
                description: 'Windows Notepad'
            },
            'paint': { 
                cmd: 'mspaint', 
                name: 'Paint',
                status: 'working',
                description: 'Microsoft Paint'
            },
            'cmd': { 
                cmd: 'cmd', 
                name: 'Command Prompt',
                status: 'working',
                description: 'Windows Command Prompt'
            },
            'powershell': { 
                cmd: 'powershell', 
                name: 'PowerShell',
                status: 'working',
                description: 'Windows PowerShell'
            },
            'explorer': { 
                cmd: 'explorer', 
                name: 'File Explorer',
                status: 'working',
                description: 'Windows File Explorer'
            },
            'taskmgr': { 
                cmd: 'taskmgr', 
                name: 'Task Manager',
                status: 'working',
                description: 'Windows Task Manager'
            },

            // EXPERIMENTAL APPS ⚠️ (may be blocked by corporate security)
            'outlook': { 
                cmd: 'outlook', 
                name: 'Microsoft Outlook',
                status: 'experimental',
                description: 'Microsoft Outlook (may be blocked by Zscaler)'
            },
            'word': { 
                cmd: 'winword', 
                name: 'Microsoft Word',
                status: 'experimental',
                description: 'Microsoft Word (may be blocked by Zscaler)'
            },
            'excel': { 
                cmd: 'excel', 
                name: 'Microsoft Excel',
                status: 'experimental',
                description: 'Microsoft Excel (may be blocked by Zscaler)'
            },
            'chrome': { 
                cmd: 'chrome', 
                name: 'Google Chrome',
                status: 'experimental',
                description: 'Google Chrome (if installed)'
            },
            'edge': { 
                cmd: 'msedge', 
                name: 'Microsoft Edge',
                status: 'experimental',
                description: 'Microsoft Edge'
            },
            'firefox': { 
                cmd: 'firefox', 
                name: 'Mozilla Firefox',
                status: 'experimental',
                description: 'Mozilla Firefox (if installed)'
            }
        };

        const app = apps[appName.toLowerCase()];
        
        if (!app) {
            resolve(`❌ Unknown application: ${appName}

✅ WORKING APPS:
• calculator, notepad, paint
• cmd, powershell, explorer, taskmgr

⚠️ EXPERIMENTAL APPS:
• outlook, word, excel, chrome, edge, firefox
(May be blocked by corporate security)`);
            return;
        }

        console.log(`🚀 Attempting to launch ${app.name} (${app.status})...`);

        const command = `start "" "${app.cmd}"`;
        
        exec(command, (error, stdout, stderr) => {
            if (error) {
                if (error.message.includes('cannot find')) {
                    if (app.status === 'working') {
                        resolve(`❌ ${app.name} not found on system
                        
This should be a standard Windows app. Possible issues:
• Windows installation incomplete
• Corporate policy restrictions
• Try opening manually from Start menu`);
                    } else {
                        resolve(`⚠️ ${app.name} not found or blocked
                        
This is an experimental feature. Possible causes:
• Application not installed
• Zscaler/corporate security blocking
• Path not in system environment

💡 Manual alternatives:
1. Press Win key
2. Type "${app.name}"
3. Click to open manually

🔧 Or try the working apps instead:
Calculator, Notepad, Paint, Command Prompt`);
                    }
                } else {
                    resolve(`❌ Launch failed: ${error.message}

💡 Try manual launch:
1. Press Windows key + R
2. Type "${app.cmd}"
3. Press Enter`);
                }
            } else {
                if (app.status === 'working') {
                    resolve(`✅ ${app.name} launched successfully!
                    
${app.description} should now be open.`);
                } else {
                    resolve(`🎯 ${app.name} launch attempted
                    
Status: ${app.status.toUpperCase()}
Description: ${app.description}

If ${app.name} didn't open, it may be:
• Blocked by corporate security (Zscaler)
• Not installed on this system
• Requiring manual launch

💡 Try opening manually from Start menu if needed.`);
                }
            }
        });
    });
}

module.exports = {
    handleAppLaunch,
    launchApplication
};