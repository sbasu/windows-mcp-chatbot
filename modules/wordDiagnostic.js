// modules/wordDiagnostic.js - WORD ISSUE DIAGNOSTIC TOOL
const { exec } = require('child_process');
const path = require('path');
const os = require('os');
const fs = require('fs');

const DESKTOP_PATH = path.join(os.homedir(), 'Desktop');

// Comprehensive Word diagnostic
async function runWordDiagnostic() {
    console.log('🔧 Starting comprehensive Word diagnostic...');
    
    const results = {
        timestamp: new Date().toISOString(),
        system: {
            os: os.platform(),
            arch: os.arch(),
            release: os.release(),
            user: os.userInfo().username
        },
        tests: [],
        recommendations: []
    };

    // Test 1: Basic Word executable detection
    const test1 = await testWordExecutable();
    results.tests.push(test1);

    // Test 2: Registry check
    const test2 = await testWordRegistry();
    results.tests.push(test2);

    // Test 3: File association test
    const test3 = await testFileAssociation();
    results.tests.push(test3);

    // Test 4: COM object test with detailed error analysis
    const test4 = await testCOMObjectDetailed();
    results.tests.push(test4);

    // Test 5: PowerShell execution policy
    const test5 = await testPowerShellPolicy();
    results.tests.push(test5);

    // Test 6: User permissions
    const test6 = await testUserPermissions();
    results.tests.push(test6);

    // Test 7: Process execution test
    const test7 = await testProcessExecution();
    results.tests.push(test7);

    // Generate recommendations based on test results
    results.recommendations = generateRecommendations(results.tests);

    // Create diagnostic report
    const report = createDiagnosticReport(results);
    
    // Save diagnostic report to desktop
    const reportPath = path.join(DESKTOP_PATH, `Word_Diagnostic_Report_${Date.now()}.txt`);
    fs.writeFileSync(reportPath, report);

    return {
        success: true,
        report: report,
        reportPath: reportPath,
        results: results
    };
}

// Test 1: Word executable detection
function testWordExecutable() {
    return new Promise((resolve) => {
        exec('where winword', (error, stdout, stderr) => {
            if (!error && stdout.trim()) {
                resolve({
                    test: 'Word Executable Detection',
                    status: 'PASS',
                    result: `Found at: ${stdout.trim()}`,
                    details: 'Word executable is available in system PATH'
                });
            } else {
                resolve({
                    test: 'Word Executable Detection',
                    status: 'FAIL',
                    result: 'Word executable not found in PATH',
                    details: 'This could mean Word is not installed or not properly registered'
                });
            }
        });
    });
}

// Test 2: Registry check
function testWordRegistry() {
    return new Promise((resolve) => {
        const command = 'reg query "HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\WINWORD.EXE" /ve';
        
        exec(command, (error, stdout, stderr) => {
            if (!error && stdout.includes('REG_SZ')) {
                const match = stdout.match(/REG_SZ\s+(.+)$/m);
                if (match) {
                    resolve({
                        test: 'Registry Check',
                        status: 'PASS',
                        result: `Registry path: ${match[1].trim()}`,
                        details: 'Word is properly registered in Windows registry'
                    });
                } else {
                    resolve({
                        test: 'Registry Check',
                        status: 'PARTIAL',
                        result: 'Registry entry exists but path unclear',
                        details: 'Word appears to be registered but path extraction failed'
                    });
                }
            } else {
                resolve({
                    test: 'Registry Check',
                    status: 'FAIL',
                    result: 'No registry entry found for Word',
                    details: 'Word may not be properly installed or registered'
                });
            }
        });
    });
}

// Test 3: File association test
function testFileAssociation() {
    return new Promise((resolve) => {
        exec('assoc .docx', (error, stdout, stderr) => {
            if (!error && stdout.includes('Word')) {
                resolve({
                    test: 'File Association',
                    status: 'PASS',
                    result: stdout.trim(),
                    details: 'DOCX files are properly associated with Word'
                });
            } else {
                resolve({
                    test: 'File Association',
                    status: 'FAIL',
                    result: error ? error.message : 'No Word association found',
                    details: 'DOCX files may not open in Word automatically'
                });
            }
        });
    });
}

// Test 4: Detailed COM object test
function testCOMObjectDetailed() {
    return new Promise((resolve) => {
        const testScript = `
$ErrorActionPreference = "Stop"
try {
    Write-Host "Testing COM object creation..."
    $word = New-Object -ComObject Word.Application
    Write-Host "SUCCESS: COM object created"
    
    Write-Host "Testing basic properties..."
    $version = $word.Version
    $build = $word.Build
    Write-Host "Version: $version, Build: $build"
    
    Write-Host "Testing document creation..."
    $doc = $word.Documents.Add()
    Write-Host "SUCCESS: Document created"
    
    Write-Host "Testing visibility..."
    $word.Visible = $true
    Start-Sleep -Seconds 2
    $word.Visible = $false
    Write-Host "SUCCESS: Visibility control works"
    
    Write-Host "Cleaning up..."
    $doc.Close($false)
    $word.Quit()
    $word = $null
    
    Write-Host "COM_TEST_SUCCESS: All COM operations successful"
    
} catch {
    $errorType = $_.Exception.GetType().Name
    $errorMessage = $_.Exception.Message
    $errorHResult = $_.Exception.HResult
    
    Write-Host "COM_TEST_FAILED: $errorType"
    Write-Host "Error Message: $errorMessage"
    Write-Host "HRESULT: 0x$($errorHResult.ToString('X8'))"
    
    # Specific error analysis
    switch ($errorHResult) {
        -2147221005 { Write-Host "ANALYSIS: Invalid class string - Word COM class not registered" }
        -2147024894 { Write-Host "ANALYSIS: File not found - Word executable missing" }
        -2147023174 { Write-Host "ANALYSIS: Access denied - Insufficient permissions" }
        -2147467259 { Write-Host "ANALYSIS: Unspecified error - Generic COM failure" }
        -2147024891 { Write-Host "ANALYSIS: Access denied - Security policy blocking" }
        default { Write-Host "ANALYSIS: Unknown COM error - May be corporate security" }
    }
}`;

        const tempScript = path.join(os.tmpdir(), `word_com_detailed_${Date.now()}.ps1`);
        fs.writeFileSync(tempScript, testScript);

        exec(`powershell -ExecutionPolicy Bypass -File "${tempScript}"`, (error, stdout, stderr) => {
            try { fs.unlinkSync(tempScript); } catch (e) {}

            if (stdout.includes('COM_TEST_SUCCESS')) {
                resolve({
                    test: 'COM Object Detailed Test',
                    status: 'PASS',
                    result: 'All COM operations successful',
                    details: stdout
                });
            } else {
                resolve({
                    test: 'COM Object Detailed Test',
                    status: 'FAIL',
                    result: 'COM operations failed',
                    details: stdout + '\n' + stderr
                });
            }
        });
    });
}

// Test 5: PowerShell execution policy
function testPowerShellPolicy() {
    return new Promise((resolve) => {
        exec('powershell -Command "Get-ExecutionPolicy"', (error, stdout, stderr) => {
            if (!error) {
                const policy = stdout.trim();
                const isRestricted = policy === 'Restricted';
                
                resolve({
                    test: 'PowerShell Execution Policy',
                    status: isRestricted ? 'WARN' : 'PASS',
                    result: `Policy: ${policy}`,
                    details: isRestricted ? 'Restricted policy may prevent automation scripts' : 'Execution policy allows scripts'
                });
            } else {
                resolve({
                    test: 'PowerShell Execution Policy',
                    status: 'FAIL',
                    result: 'Could not check execution policy',
                    details: error.message
                });
            }
        });
    });
}

// Test 6: User permissions
function testUserPermissions() {
    return new Promise((resolve) => {
        exec('whoami /groups | find "S-1-16-12288"', (error, stdout, stderr) => {
            const isElevated = !error && stdout.includes('S-1-16-12288');
            
            resolve({
                test: 'User Permissions',
                status: isElevated ? 'PASS' : 'WARN',
                result: isElevated ? 'Running with elevated privileges' : 'Running with standard privileges',
                details: isElevated ? 'Administrative rights available' : 'Some operations may require elevation'
            });
        });
    });
}

// Test 7: Process execution test
function testProcessExecution() {
    return new Promise((resolve) => {
        exec('start /wait /min winword /q', (error, stdout, stderr) => {
            setTimeout(() => {
                exec('taskkill /f /im winword.exe 2>nul', () => {
                    if (!error) {
                        resolve({
                            test: 'Process Execution',
                            status: 'PASS',
                            result: 'Word process can be started',
                            details: 'Word executable launches successfully'
                        });
                    } else {
                        resolve({
                            test: 'Process Execution',
                            status: 'FAIL',
                            result: 'Cannot start Word process',
                            details: error.message
                        });
                    }
                });
            }, 3000);
        });
    });
}

// Generate recommendations based on test results
function generateRecommendations(tests) {
    const recommendations = [];
    
    const executableTest = tests.find(t => t.test === 'Word Executable Detection');
    const comTest = tests.find(t => t.test === 'COM Object Detailed Test');
    const registryTest = tests.find(t => t.test === 'Registry Check');
    const policyTest = tests.find(t => t.test === 'PowerShell Execution Policy');
    const permTest = tests.find(t => t.test === 'User Permissions');

    // Word not installed
    if (executableTest?.status === 'FAIL' && registryTest?.status === 'FAIL') {
        recommendations.push({
            priority: 'HIGH',
            issue: 'Microsoft Word not installed',
            solution: 'Install Microsoft Office or Office 365',
            steps: [
                'Visit office.com or contact IT department',
                'Install Office suite with Word application',
                'Ensure proper licensing and activation'
            ]
        });
    }

    // COM blocked by security
    if (executableTest?.status === 'PASS' && comTest?.status === 'FAIL') {
        if (comTest.details?.includes('Access denied') || comTest.details?.includes('Security policy')) {
            recommendations.push({
                priority: 'HIGH',
                issue: 'COM automation blocked by corporate security',
                solution: 'Use file-based alternatives instead of COM automation',
                steps: [
                    'Use RTF file creation instead of direct Word automation',
                    'Create documents and open them manually in Word',
                    'Contact IT about COM automation policies if needed'
                ]
            });
        }
    }

    // PowerShell restricted
    if (policyTest?.status === 'WARN') {
        recommendations.push({
            priority: 'MEDIUM',
            issue: 'PowerShell execution policy is restricted',
            solution: 'Adjust PowerShell execution policy',
            steps: [
                'Run PowerShell as Administrator',
                'Execute: Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser',
                'Or use the -ExecutionPolicy Bypass parameter in scripts'
            ]
        });
    }

    // Permissions
    if (permTest?.status === 'WARN') {
        recommendations.push({
            priority: 'LOW',
            issue: 'Running without elevated privileges',
            solution: 'Try running as administrator for some operations',
            steps: [
                'Right-click Command Prompt → Run as Administrator',
                'Run the application with elevated privileges',
                'Note: This may not be possible in corporate environments'
            ]
        });
    }

    // Default recommendation if no specific issues found
    if (recommendations.length === 0) {
        recommendations.push({
            priority: 'MEDIUM',
            issue: 'Word automation not working despite tests passing',
            solution: 'Use alternative document creation methods',
            steps: [
                'Use RTF file creation (works without COM)',
                'Create HTML documents that Word can open',
                'Use text templates for manual copy/paste',
                'Consider using Google Docs or online alternatives'
            ]
        });
    }

    return recommendations;
}

// Create diagnostic report
function createDiagnosticReport(results) {
    const report = `MICROSOFT WORD DIAGNOSTIC REPORT
${'='.repeat(50)}

Generated: ${new Date().toLocaleString()}
System: ${results.system.os} ${results.system.arch} (${results.system.release})
User: ${results.system.user}

${'='.repeat(50)}
TEST RESULTS
${'='.repeat(50)}

${results.tests.map(test => `
${test.test}:
Status: ${test.status}
Result: ${test.result}
Details: ${test.details}
`).join('\n')}

${'='.repeat(50)}
RECOMMENDATIONS
${'='.repeat(50)}

${results.recommendations.map((rec, index) => `
${index + 1}. ${rec.issue} [${rec.priority} PRIORITY]
Solution: ${rec.solution}
Steps:
${rec.steps.map(step => `   • ${step}`).join('\n')}
`).join('\n')}

${'='.repeat(50)}
ALTERNATIVE SOLUTIONS
${'='.repeat(50)}

If Word automation continues to fail, use these alternatives:

1. RTF FILE METHOD (Recommended)
   • Creates Rich Text Format files
   • Double-click to open in Word
   • Preserves formatting and structure
   • Works without COM automation

2. HTML FILE METHOD
   • Creates HTML documents
   • Open Word → File → Open → Select HTML file
   • Good formatting compatibility
   • Cross-platform compatible

3. TEXT TEMPLATE METHOD
   • Creates formatted text templates
   • Copy and paste into Word manually
   • Always works regardless of restrictions
   • Good for troubleshooting

4. ONLINE ALTERNATIVES
   • Google Docs (docs.google.com)
   • Microsoft Office Online (office.com)
   • Works from any browser
   • No local software requirements

${'='.repeat(50)}
TECHNICAL NOTES
${'='.repeat(50)}

Common corporate restrictions:
• Zscaler/proxy blocking COM objects
• Group Policy preventing automation
• Antivirus blocking script execution
• User Account Control (UAC) restrictions
• Office license/activation issues

Contact your IT department if you need:
• COM automation policy exceptions
• PowerShell execution policy changes
• Administrative privileges
• Office installation/licensing support

${'='.repeat(50)}
END REPORT
${'='.repeat(50)}`;

    return report;
}

// Simple Word alternative - guaranteed to work
async function createSimpleDocument(content, title = 'Document') {
    try {
        const timestamp = Date.now();
        const results = [];

        // 1. Create simple RTF
        const rtfContent = `{\\rtf1\\ansi\\deff0 {\\fonttbl {\\f0 Times New Roman;}}
\\f0\\fs24 {\\b\\fs28 ${title}\\b0\\fs24\\par\\par}
${content.replace(/\n/g, '\\par ')}
}`;
        
        const rtfPath = path.join(DESKTOP_PATH, `${title.replace(/\s/g, '_')}_${timestamp}.rtf`);
        fs.writeFileSync(rtfPath, rtfContent);
        results.push(`📄 ${path.basename(rtfPath)} - RTF format`);

        // 2. Create simple text file
        const txtContent = `${title}\n${'='.repeat(title.length)}\n\n${content}\n\nCreated: ${new Date().toLocaleString()}`;
        const txtPath = path.join(DESKTOP_PATH, `${title.replace(/\s/g, '_')}_${timestamp}.txt`);
        fs.writeFileSync(txtPath, txtContent);
        results.push(`📝 ${path.basename(txtPath)} - Text format`);

        return `✅ SIMPLE DOCUMENT CREATED:

📁 Files on Desktop:
${results.map(r => '• ' + r).join('\n')}

💡 TO OPEN IN WORD:
1. Double-click the .rtf file (should open in Word)
2. If that fails, open Word manually
3. File → Open → Select the .rtf file
4. Or copy text from .txt file and paste into Word

This method bypasses all automation and should always work!`;

    } catch (error) {
        return `❌ Simple document creation failed: ${error.message}

Try manual creation:
1. Open Notepad
2. Type your content
3. Save as .txt file
4. Copy content and paste into Word`;
    }
}

module.exports = {
    runWordDiagnostic,
    createSimpleDocument
};