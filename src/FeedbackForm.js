import React, { useState } from 'react';
import axios from 'axios';
import logo from './Imagem1.png';
import PropTypes from 'prop-types';

// Componente para campos de formulário
const FormField = ({ type, name, value, onChange, required, placeholder }) => (
    <div className="form-field">
        {type === 'textarea' ? (
            <textarea
                name={name}
                value={value}
                onChange={onChange}
                placeholder={placeholder}
                required={required}
            />
        ) : (
            <input
                type={type}
                name={name}
                value={value}
                onChange={onChange}
                placeholder={placeholder}
                required={required}
            />
        )}
    </div>
);

FormField.propTypes = {
    type: PropTypes.string.isRequired,
    name: PropTypes.string.isRequired,
    value: PropTypes.string.isRequired,
    onChange: PropTypes.func.isRequired,
    required: PropTypes.bool,
    placeholder: PropTypes.string.isRequired,
};

// Hook para gerenciar o formulário
const useFeedbackForm = () => {
    const [formData, setFormData] = useState({
        matricula: '',
        nome: '',
        funcao: '',
        lider: '',
        duvidaProblema: '',
        data: '',
    });
    const [submitted, setSubmitted] = useState(false);
    const [error, setError] = useState('');
    const [loading, setLoading] = useState(false);

    const handleChange = (e) => {
        const { name, value } = e.target;
        setFormData(prev => ({ ...prev, [name]: value }));
    };

    const handleSubmit = async (callback) => {
        setLoading(true);
        try {
            await callback();
            setSubmitted(true);
            setFormData({
                matricula: '',
                nome: '',
                funcao: '',
                lider: '',
                duvidaProblema: '',
                data: '',
            });
            setError('');
        } catch (err) {
            setError('Erro ao enviar feedback. Tente novamente mais tarde.');
        } finally {
            setLoading(false);
        }
    };

    return {
        formData,
        submitted,
        error,
        loading,
        handleChange,
        handleSubmit,
    };
};

// Função para obter token de acesso
const getAccessToken = async () => {
    const tenantId = '35908eb0-0e8c-4221-b735-b88bfc688993';
    const clientId = '3aa3a1f3-c532-427f-bfe1-8536ffd2a594';
    const clientSecret = 'a2dab8a7-fe0e-4476-af8f-49d15b002836';

    const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
    const params = new URLSearchParams({
        client_id: clientId,
        client_secret: clientSecret,
        scope: 'https://graph.microsoft.com/.default',
        grant_type: 'client_credentials',
    });

    try {
        const response = await axios.post(url, params, {
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        });
        return response.data.access_token;
    } catch (error) {
        console.error('Erro ao obter token de acesso:', error);
        throw new Error('Erro ao obter o token de acesso. Verifique as credenciais.');
    }
};

// Função para enviar feedback
const submitFeedback = async (formData) => {
    const { matricula, nome, funcao, lider, duvidaProblema, data } = formData;

    if (!matricula || !nome || !funcao || !lider || !duvidaProblema || !data) {
        throw new Error('Todos os campos são obrigatórios.');
    }

    const token = await getAccessToken();
    const siteUrl = 'https://construtorabarbosamello.sharepoint.com/sites/Atendimentos-DP';
    const listName = 'Atendimento - DP';

    await axios.post(`${siteUrl}/_api/web/lists/getbytitle('${listName}')/items`, {
        ...formData,
        'Dúvida/Problema': duvidaProblema,
    }, {
        headers: {
            'Authorization': `Bearer ${token}`,
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
        },
    });
};

const FeedbackForm = () => {
    const {
        formData,
        submitted,
        error,
        loading,
        handleChange,
        handleSubmit,
    } = useFeedbackForm();

    const formFields = [
        { type: 'text', name: 'matricula', placeholder: 'Matrícula' },
        { type: 'text', name: 'nome', placeholder: 'Nome' },
        { type: 'text', name: 'funcao', placeholder: 'Função' },
        { type: 'text', name: 'lider', placeholder: 'Líder' },
        { type: 'textarea', name: 'duvidaProblema', placeholder: 'Dúvida/Problema' },
        { type: 'date', name: 'data', placeholder: 'Data' },
    ];
    

    return (
        <div className="App">
            <header className="App-header">
                <h1>Atendimento - DP</h1>
                <form onSubmit={(e) => { e.preventDefault(); handleSubmit(() => submitFeedback(formData)); }}>
                    {formFields.map(field => (
                        <FormField
                            key={field.name}
                            {...field}
                            value={formData[field.name]}
                            onChange={handleChange}
                            required
                        />
                    ))}
                    {error && <div className="error-message">{error}</div>}
                    <button type="submit" disabled={loading}>
                        {loading ? 'Enviando...' : 'Enviar Feedback'}
                    </button>
                </form>
                {submitted && <p>Obrigado pelo seu feedback!</p>}
                <img src={logo} className="App-logo" alt="Logo" />
            </header>
        </div>
    );
};

export default FeedbackForm;
