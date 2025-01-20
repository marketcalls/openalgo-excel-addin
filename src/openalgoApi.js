export class OpenAlgoAPI {
    constructor(host = 'http://127.0.0.1:5000', apikey = '') {
        this.host = host;
        this.apikey = apikey;
    }

    async getFunds() {
        try {
            const response = await fetch(`${this.host}/api/v1/funds`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    apikey: this.apikey
                })
            });

            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }

            const data = await response.json();
            return data;
        } catch (error) {
            console.error('Error fetching funds:', error);
            throw error;
        }
    }
}
