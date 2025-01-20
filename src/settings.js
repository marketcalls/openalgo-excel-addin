export class Settings {
    static async saveSettings(host, apikey) {
        try {
            await Office.context.document.settings.set('openalgo_host', host);
            await Office.context.document.settings.set('openalgo_apikey', apikey);
            await Office.context.document.settings.saveAsync();
            return true;
        } catch (error) {
            console.error('Error saving settings:', error);
            return false;
        }
    }

    static getSettings() {
        const host = Office.context.document.settings.get('openalgo_host') || 'http://127.0.0.1:5000';
        const apikey = Office.context.document.settings.get('openalgo_apikey') || '';
        return { host, apikey };
    }
}
