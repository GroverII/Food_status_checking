// api.js
const baseUrl = 'https://localhost:7045/SchoolFood';

export const getAllClasses = async () => {
    try {
        const response = await fetch(`${baseUrl}/get-all-classes`);

        if (!response.ok) {
            throw new Error(`Request failed: ${response.url} - ${response.status} ${response.statusText}`);
        }

        const data = await response.json();
        return data;
    } catch (error) {
        console.error('Error fetching classes:', error.message);
        throw error;
    }
};

export const getStudentData = async (requestBody) => {
    try {
        const response = await fetch(`${baseUrl}/get-student-data`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(requestBody),
        });

        if (!response.ok) {
            throw new Error(`Request failed: ${response.url} - ${response.status} ${response.statusText}`);
        }

        const excelBlob = await response.blob();
        return excelBlob;
    } catch (error) {
        console.error('Error fetching student data:', error.message);
        throw error;
    }
};

export const findStudent = async (codeValue) => {
    try {
        const response = await fetch(`${baseUrl}/find/${codeValue}`, {
            method: 'GET',
            headers: {
                'Content-Type': 'application/json',
            },
        });

        if (!response.ok) {
            throw new Error(`Request failed: ${response.url} - ${response.status} ${response.statusText}`);
        }

        const result = await response.text();
        return result;
    } catch (error) {
        console.error('Error finding data:', error.message);
        throw error;
    }
};
