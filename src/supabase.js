const API_BASE_URL = 'http://localhost:3000/api';

const auth = {
    async signUp(email, password, userData) {
        try {
            const response = await fetch(`${API_BASE_URL}/auth/signup`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    email,
                    password,
                    name: userData.name
                })
            });

            const result = await response.json();
            return result;
        } catch (error) {
            return { data: null, error: error.message };
        }
    },

    async initializeUserCredits(userId) {
        try {
            const response = await fetch(`${API_BASE_URL}/auth/initialize-credits`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    userId,
                    initialCredits: 20
                })
            });

            const result = await response.json();
            return result;
        } catch (error) {
            return { error: error.message };
        }
    },

    async signIn(email, password) {
        try {
            const response = await fetch(`${API_BASE_URL}/auth/signin`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    email,
                    password
                })
            });

            const result = await response.json();
            return result;
        } catch (error) {
            return { data: null, error: error.message };
        }
    },

    async signOut() {
        try {
            const response = await fetch(`${API_BASE_URL}/auth/signout`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                }
            });

            const result = await response.json();
            return result;
        } catch (error) {
            return { error: error.message };
        }
    }
};

export default auth;