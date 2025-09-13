// server.test.js
const { commands } = require('./server');

describe('Chatbot Commands', () => {
    afterEach(() => {
        jest.restoreAllMocks();
    });

    it('should return a default message for an unknown command', async () => {
        const response = await commands.processUserMessage('unknown command');
        expect(response.message).toBe('🤖 Simple System Ready!');
    });

    it('should handle the "screenshot" command', async () => {
        const takeSmartScreenshotSpy = jest.spyOn(commands, 'takeSmartScreenshot').mockResolvedValue({ success: true, message: 'Screenshot taken' });
        const response = await commands.processUserMessage('take a screenshot');
        expect(response.message).toBe('📸 Screenshot taken!');
        expect(takeSmartScreenshotSpy).toHaveBeenCalled();
    });

    it('should handle the "list screenshots" command', async () => {
        const listRecentScreenshotsSpy = jest.spyOn(commands, 'listRecentScreenshots').mockResolvedValue('screenshot1.png');
        const response = await commands.processUserMessage('list screenshots');
        expect(response.message).toBe('📁 Recent screenshots:');
        expect(listRecentScreenshotsSpy).toHaveBeenCalled();
    });

    it('should handle the "open" command', async () => {
        const launchSimpleAppSpy = jest.spyOn(commands, 'launchSimpleApp').mockResolvedValue({ success: true });
        const response = await commands.processUserMessage('open calculator');
        expect(response.message).toBe('🚀 calculator launched!');
        expect(launchSimpleAppSpy).toHaveBeenCalledWith('calculator');
    });

    it('should handle the "system info" command', async () => {
        const getSimpleSystemInfoSpy = jest.spyOn(commands, 'getSimpleSystemInfo').mockResolvedValue('Windows 10 Pro');
        const response = await commands.processUserMessage('system info');
        expect(response.message).toBe('💻 System information:');
        expect(getSimpleSystemInfoSpy).toHaveBeenCalled();
    });

    it('should handle errors gracefully', async () => {
        const getSimpleSystemInfoSpy = jest.spyOn(commands, 'getSimpleSystemInfo').mockRejectedValue(new Error('test error'));
        const response = await commands.processUserMessage('system info');
        expect(response.message).toBe('❌ Command failed');
    });
});
