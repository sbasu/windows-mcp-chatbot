// modules/systemInfo.js - SYSTEM INFO MODULE (WORKING)
const { exec } = require('child_process');

async function handleSystemInfo() {
    try {
        const info = await getSystemInformation();
        
        return {
            message: '💻 System Information Retrieved',
            details: info
        };
    } catch (error) {
        return {
            message: '❌ System info error',
            details: error.message
        };
    }
}

function getSystemInformation() {
    return new Promise((resolve) => {
        // Use multiple commands to get comprehensive system info
        const commands = [
            'systeminfo | findstr /C:"OS Name" /C:"OS Version" /C:"System Type" /C:"Total Physical Memory" /C:"System Manufacturer" /C:"System Model"',
            'wmic cpu get name /value',
            'wmic diskdrive get size,model /value'
        ];

        let results = [];
        let completed = 0;

        // Execute system info command
        exec(commands[0], (error, stdout, stderr) => {
            if (!error && stdout) {
                results.push('🖥️ SYSTEM INFORMATION:\n' + stdout.trim());
            } else {
                results.push('⚠️ Basic system info not available');
            }
            
            completed++;
            if (completed === commands.length) {
                resolve(formatSystemInfo(results));
            }
        });

        // Execute CPU info command
        exec(commands[1], (error, stdout, stderr) => {
            if (!error && stdout) {
                const cpuMatch = stdout.match(/Name=(.+)/);
                if (cpuMatch) {
                    results.push('\n🔧 PROCESSOR:\n' + cpuMatch[1].trim());
                }
            }
            
            completed++;
            if (completed === commands.length) {
                resolve(formatSystemInfo(results));
            }
        });

        // Execute disk info command  
        exec(commands[2], (error, stdout, stderr) => {
            if (!error && stdout) {
                const sizeMatches = stdout.match(/Size=(\d+)/g);
                const modelMatches = stdout.match(/Model=(.+)/g);
                
                if (sizeMatches && modelMatches) {
                    results.push('\n💾 STORAGE:');
                    sizeMatches.forEach((size, index) => {
                        const sizeGB = Math.round(parseInt(size.split('=')[1]) / (1024 * 1024 * 1024));
                        const model = modelMatches[index] ? modelMatches[index].split('=')[1].trim() : 'Unknown';
                        results.push(`• ${model}: ${sizeGB} GB`);
                    });
                }
            }
            
            completed++;
            if (completed === commands.length) {
                resolve(formatSystemInfo(results));
            }
        });

        // Fallback timeout
        setTimeout(() => {
            if (completed < commands.length) {
                resolve(results.length > 0 ? formatSystemInfo(results) : 'System information partially available');
            }
        }, 5000);
    });
}

function formatSystemInfo(results) {
    if (results.length === 0) {
        return '❌ Unable to retrieve system information\n\nThis may be due to:\n• Corporate security restrictions\n• Insufficient permissions\n• System command limitations';
    }

    let formatted = results.join('\n');
    
    // Add current date/time
    const now = new Date();
    formatted = `📊 SYSTEM REPORT - ${now.toLocaleDateString()} ${now.toLocaleTimeString()}\n${'='.repeat(60)}\n\n${formatted}`;
    
    // Add helpful footer
    formatted += `\n\n${'='.repeat(60)}
💡 ADDITIONAL INFO:
• Node.js Version: ${process.version}
• Platform: ${process.platform}
• Architecture: ${process.arch}
• Uptime: ${Math.floor(process.uptime())} seconds

🔧 For more detailed info, try:
• Task Manager → Performance tab
• Settings → System → About
• Control Panel → System`;

    return formatted;
}

// Additional utility function for quick system check
function getQuickSystemInfo() {
    return new Promise((resolve) => {
        const info = {
            platform: process.platform,
            arch: process.arch,
            nodeVersion: process.version,
            uptime: process.uptime(),
            memory: process.memoryUsage(),
            cpu: require('os').cpus().length + ' cores',
            hostname: require('os').hostname(),
            userInfo: require('os').userInfo().username
        };

        resolve(`💻 QUICK SYSTEM INFO:
Platform: ${info.platform}
Architecture: ${info.arch}
Node.js: ${info.nodeVersion}
CPU Cores: ${info.cpu}
Hostname: ${info.hostname}
User: ${info.userInfo}
Uptime: ${Math.floor(info.uptime)} seconds`);
    });
}

module.exports = {
    handleSystemInfo,
    getSystemInformation,
    getQuickSystemInfo
};