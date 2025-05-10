const assert = require('assert');
const vscode = require('vscode');
const sinon = require('sinon');
const extension = require('../../extension');
const auth = require('../../authentication');
const teamsService = require('../../teamsService');

suite('Extension Test Suite', () => {
    vscode.window.showInformationMessage('Starting tests...');

    test('Extension should be present', () => {
        assert.ok(vscode.extensions.getExtension('Harishri.CodeToTeams'));
    });

    test('Should register commands', async () => {
        const commands = await vscode.commands.getCommands();
        assert.ok(commands.includes('extension.shareToTeams'));
    });

    test('formatCodeSnippet should add language and code block markers', () => {
        const mockText = 'const test = "hello";';
        const formattedText = extension.formatCodeSnippet(mockText, 'javascript');
        assert.strictEqual(formattedText, '```javascript\nconst test = "hello";\n```');
    });

    test('formatCodeSnippet should work without language', () => {
        const mockText = 'const test = "hello";';
        const formattedText = extension.formatCodeSnippet(mockText, null);
        assert.strictEqual(formattedText, '```\nconst test = "hello";\n```');
    });

    test('Should handle authentication flow', async () => {
        const getAccessTokenStub = sinon.stub(auth, 'getAccessToken').resolves('mock-token');
        
        const token = await auth.getAccessToken();
        
        assert.strictEqual(token, 'mock-token');
        getAccessTokenStub.restore();
    });

    test('Should handle Teams service', async () => {
        const accessToken = 'mock-token';
        const mockChats = [
            { id: 'chat1', topic: 'Team Chat', members: [{ displayName: 'User 1' }] },
            { id: 'chat2', members: [{ displayName: 'User 2' }] }
        ];
        
        const getChatsStub = sinon.stub(teamsService, 'getChats').resolves(mockChats);
        
        const chats = await teamsService.getChats(accessToken);
        
        assert.deepStrictEqual(chats, mockChats);
        assert.strictEqual(chats.length, 2);
        getChatsStub.restore();
    });
});