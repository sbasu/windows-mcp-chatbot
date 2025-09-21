// Test script to check module imports after fixes
console.log('🔧 TESTING MODULE IMPORTS AFTER FIXES...');
console.log('==========================================');

try {
    console.log('\n📸 Testing screenshot module...');
    const screenshotModule = require('./modules/screenshot');
    console.log('✅ Screenshot module loaded successfully');
    
    console.log('\n🚀 Testing appLauncher module...');
    const appLauncherModule = require('./modules/appLauncher');
    console.log('✅ AppLauncher module loaded successfully');
    
    console.log('\n💻 Testing systeminfo module (FIXED IMPORT PATH)...');
    const systemInfoModule = require('./modules/systeminfo');
    console.log('✅ SystemInfo module loaded successfully - IMPORT ISSUE FIXED!');
    
    console.log('\n📧 Testing email module...');
    const emailModule = require('./modules/email');
    console.log('✅ Email module loaded successfully');
    
    console.log('\n🤖 Testing claude module...');
    const claudeModule = require('./modules/claude');
    console.log('✅ Claude module loaded successfully');
    
    console.log('\n📝 Testing wordSimple module (FIXED CORRUPTED FILE)...');
    const wordSimple = require('./modules/wordSimple');
    console.log('✅ WordSimple module loaded successfully - CORRUPTED FILE FIXED!');
    
    console.log('\n🔧 Testing wordDiagnostic module...');
    const wordDiagnostic = require('./modules/wordDiagnostic');
    console.log('✅ WordDiagnostic module loaded successfully');
    
    console.log('\n🎉 ==========================================');
    console.log('🎉 ALL MODULES IMPORTED SUCCESSFULLY!');
    console.log('🎉 ==========================================');
    console.log('\n✅ PRIMARY FIXES VERIFIED:');
    console.log('   • systeminfo.js import path corrected');
    console.log('   • wordSimple.js file content restored');
    console.log('   • All 7 modules loading without errors');
    console.log('\n🚀 READY TO START SERVER: node server.js');
    
} catch (error) {
    console.error('\n❌ ==========================================');
    console.error('❌ MODULE IMPORT STILL FAILING!');
    console.error('❌ ==========================================');
    console.error('❌ Error details:', error.message);
    console.error('❌ Stack trace:', error.stack);
    console.error('\n🔧 TROUBLESHOOTING STEPS:');
    console.error('   1. Check file exists: ls -la modules/');
    console.error('   2. Check exact filename spelling');
    console.error('   3. Verify file permissions');
    console.error('   4. Check for syntax errors in the module');
}
