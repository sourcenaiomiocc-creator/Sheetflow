// Inicialização do Firebase (sem módulos ES6)

// Configuração do Firebase
const firebaseConfig = {
  apiKey: "AIzaSyDMmzPxu-Z1VEuY5WJzme1qZ5V_rqTmjaE",
  authDomain: "planilha-app-ac883.firebaseapp.com",
  projectId: "planilha-app-ac883",
  storageBucket: "planilha-app-ac883.firebasestorage.app",
  messagingSenderId: "900768678386",
  appId: "1:900768678386:web:69561ff7a05ea04e0c2b6b",
  measurementId: "G-H53P0130MF"
};

// Inicializar Firebase
firebase.initializeApp(firebaseConfig);

// Referências globais
const auth = firebase.auth();
const db = firebase.firestore();
const googleProvider = new firebase.auth.GoogleAuthProvider();

// Estado do usuário atual
let currentUser = null;

// Listener de autenticação
auth.onAuthStateChanged(async (user) => {
    currentUser = user;

    if (user) {
        console.log('Usuário logado:', user.displayName);

        // Salvar/atualizar informações do usuário no Firestore
        try {
            await db.collection('users').doc(user.uid).set({
                uid: user.uid,
                displayName: user.displayName,
                email: user.email,
                photoURL: user.photoURL,
                lastLogin: firebase.firestore.FieldValue.serverTimestamp()
            }, { merge: true });
        } catch (error) {
            console.error('Erro ao salvar usuário:', error);
        }

        showUserInterface(user);
        await loadUserSpreadsheets();
        renderSavedSpreadsheets();
    } else {
        console.log('Usuário não logado');
        showLoginInterface();
    }
});

// Login com Google
async function loginWithGoogle() {
    try {
        const result = await auth.signInWithPopup(googleProvider);
        const user = result.user;
        showNotification(`Bem-vindo, ${user.displayName}!`, 'success');
        return user;
    } catch (error) {
        console.error('Erro no login:', error);
        showNotification('Erro ao fazer login: ' + error.message, 'error');
        throw error;
    }
}

// Logout
async function logout() {
    try {
        await auth.signOut();
        showNotification('Logout realizado com sucesso!', 'success');
        window.location.reload();
    } catch (error) {
        console.error('Erro no logout:', error);
        showNotification('Erro ao fazer logout: ' + error.message, 'error');
    }
}

// Salvar planilha no Firestore
async function saveSpreadsheetToFirestore(spreadsheet) {
    if (!currentUser) {
        showNotification('Você precisa estar logado para salvar!', 'error');
        return;
    }

    try {
        // Converter array 2D para formato aceito pelo Firestore
        const spreadsheetData = {
            ...spreadsheet,
            // Converter matriz 2D em JSON string (Firestore não aceita arrays aninhados)
            data: JSON.stringify(spreadsheet.data),
            owner: currentUser.uid,
            ownerName: currentUser.displayName,
            ownerEmail: currentUser.email,
            updatedAt: firebase.firestore.FieldValue.serverTimestamp()
        };

        // Inicializar permissões se não existir
        if (!spreadsheetData.permissions) {
            spreadsheetData.permissions = {};
        }
        spreadsheetData.permissions[currentUser.uid] = 'owner';

        if (!spreadsheet.firestoreId) {
            // Criar novo documento
            spreadsheetData.createdAt = firebase.firestore.FieldValue.serverTimestamp();
            const docRef = await db.collection('spreadsheets').add(spreadsheetData);

            showNotification('Planilha salva no Firebase!', 'success');
            return docRef.id;
        } else {
            // Atualizar documento existente
            await db.collection('spreadsheets').doc(spreadsheet.firestoreId).update(spreadsheetData);
            showNotification('Planilha atualizada no Firebase!', 'success');
            return spreadsheet.firestoreId;
        }
    } catch (error) {
        console.error('Erro ao salvar planilha:', error);
        showNotification('Erro ao salvar: ' + error.message, 'error');
        throw error;
    }
}

// Carregar planilhas do usuário
async function loadUserSpreadsheets() {
    if (!currentUser) return [];

    try {
        const querySnapshot = await db.collection('spreadsheets')
            .where(`permissions.${currentUser.uid}`, 'in', ['owner', 'editor', 'viewer'])
            .orderBy('updatedAt', 'desc')
            .get();

        spreadsheets = [];
        querySnapshot.forEach((doc) => {
            const data = doc.data();
            spreadsheets.push({
                firestoreId: doc.id,
                ...data,
                // Converter JSON string de volta para array 2D
                data: typeof data.data === 'string' ? JSON.parse(data.data) : data.data
            });
        });

        return spreadsheets;
    } catch (error) {
        console.error('Erro ao carregar planilhas:', error);

        // Se der erro de índice, tentar sem orderBy
        try {
            const querySnapshot = await db.collection('spreadsheets').get();
            spreadsheets = [];

            querySnapshot.forEach((doc) => {
                const data = doc.data();
                // Filtrar manualmente
                if (data.permissions && data.permissions[currentUser.uid]) {
                    spreadsheets.push({
                        firestoreId: doc.id,
                        ...data,
                        // Converter JSON string de volta para array 2D
                        data: typeof data.data === 'string' ? JSON.parse(data.data) : data.data
                    });
                }
            });

            // Ordenar manualmente
            spreadsheets.sort((a, b) => {
                const dateA = a.updatedAt?.toDate ? a.updatedAt.toDate() : new Date(a.updatedAt);
                const dateB = b.updatedAt?.toDate ? b.updatedAt.toDate() : new Date(b.updatedAt);
                return dateB - dateA;
            });

            return spreadsheets;
        } catch (error2) {
            console.error('Erro ao carregar planilhas (fallback):', error2);
            showNotification('Erro ao carregar planilhas. Verifique as regras do Firestore.', 'error');
            return [];
        }
    }
}

// Carregar planilha específica
async function loadSpreadsheetFromFirestore(firestoreId) {
    if (!currentUser) {
        showNotification('Você precisa estar logado!', 'error');
        return null;
    }

    try {
        const doc = await db.collection('spreadsheets').doc(firestoreId).get();

        if (doc.exists) {
            const data = doc.data();

            // Verificar permissão
            if (!data.permissions || !data.permissions[currentUser.uid]) {
                showNotification('Você não tem permissão para acessar esta planilha', 'error');
                return null;
            }

            // Converter JSON string de volta para array 2D
            const spreadsheet = {
                firestoreId: doc.id,
                ...data,
                data: typeof data.data === 'string' ? JSON.parse(data.data) : data.data
            };

            return spreadsheet;
        } else {
            showNotification('Planilha não encontrada', 'error');
            return null;
        }
    } catch (error) {
        console.error('Erro ao carregar planilha:', error);
        showNotification('Erro ao carregar: ' + error.message, 'error');
        return null;
    }
}

// Excluir planilha
async function deleteSpreadsheetFromFirestore(firestoreId) {
    if (!currentUser) return false;

    try {
        const doc = await db.collection('spreadsheets').doc(firestoreId).get();

        if (doc.exists) {
            const data = doc.data();

            // Apenas o owner pode excluir
            if (data.owner !== currentUser.uid) {
                showNotification('Apenas o dono pode excluir esta planilha', 'error');
                return false;
            }

            await db.collection('spreadsheets').doc(firestoreId).delete();
            showNotification('Planilha excluída!', 'success');
            return true;
        }
    } catch (error) {
        console.error('Erro ao excluir:', error);
        showNotification('Erro ao excluir: ' + error.message, 'error');
        return false;
    }
}

// Compartilhar planilha com outro usuário
async function shareSpreadsheet(firestoreId, userEmail, permission = 'viewer') {
    if (!currentUser) return false;

    try {
        const doc = await db.collection('spreadsheets').doc(firestoreId).get();

        if (!doc.exists) {
            showNotification('Planilha não encontrada', 'error');
            return false;
        }

        const data = doc.data();

        // Verificar se o usuário atual é owner ou editor
        if (data.permissions[currentUser.uid] !== 'owner' && data.permissions[currentUser.uid] !== 'editor') {
            showNotification('Você não tem permissão para compartilhar esta planilha', 'error');
            return false;
        }

        // Buscar usuário por email
        const usersSnapshot = await db.collection('users').where('email', '==', userEmail).get();

        if (usersSnapshot.empty) {
            showNotification('Usuário não encontrado. Ele precisa fazer login no sistema primeiro.', 'warning');
            return false;
        }

        const targetUser = usersSnapshot.docs[0].data();

        // Atualizar permissões
        const updateData = {
            updatedAt: firebase.firestore.FieldValue.serverTimestamp()
        };
        updateData[`permissions.${targetUser.uid}`] = permission;

        await db.collection('spreadsheets').doc(firestoreId).update(updateData);

        // Criar notificação
        await db.collection('notifications').add({
            userId: targetUser.uid,
            type: 'share',
            spreadsheetId: firestoreId,
            spreadsheetName: data.name,
            sharedBy: currentUser.displayName,
            sharedByEmail: currentUser.email,
            permission: permission,
            read: false,
            createdAt: firebase.firestore.FieldValue.serverTimestamp()
        });

        showNotification(`Planilha compartilhada com ${userEmail} (${permission})!`, 'success');
        return true;
    } catch (error) {
        console.error('Erro ao compartilhar:', error);
        showNotification('Erro ao compartilhar: ' + error.message, 'error');
        return false;
    }
}

// Remover acesso de usuário
async function removeUserAccess(firestoreId, userId) {
    if (!currentUser) return false;

    try {
        const doc = await db.collection('spreadsheets').doc(firestoreId).get();

        if (!doc.exists) return false;

        const data = doc.data();

        // Apenas owner pode remover acessos
        if (data.owner !== currentUser.uid) {
            showNotification('Apenas o dono pode remover acessos', 'error');
            return false;
        }

        if (userId === currentUser.uid) {
            showNotification('Você não pode remover seu próprio acesso', 'error');
            return false;
        }

        // Remover permissão
        const updateData = {
            updatedAt: firebase.firestore.FieldValue.serverTimestamp()
        };
        updateData[`permissions.${userId}`] = firebase.firestore.FieldValue.delete();

        await db.collection('spreadsheets').doc(firestoreId).update(updateData);
        showNotification('Acesso removido!', 'success');
        return true;
    } catch (error) {
        console.error('Erro ao remover acesso:', error);
        showNotification('Erro ao remover acesso: ' + error.message, 'error');
        return false;
    }
}

// Obter lista de usuários com acesso
async function getSpreadsheetCollaborators(firestoreId) {
    if (!currentUser) return [];

    try {
        const doc = await db.collection('spreadsheets').doc(firestoreId).get();

        if (!doc.exists) return [];

        const data = doc.data();
        const permissions = data.permissions || {};
        const collaborators = [];

        // Buscar informações de cada usuário
        for (const [userId, permission] of Object.entries(permissions)) {
            if (permission) {
                const userDoc = await db.collection('users').doc(userId).get();
                if (userDoc.exists) {
                    collaborators.push({
                        ...userDoc.data(),
                        permission: permission
                    });
                }
            }
        }

        return collaborators;
    } catch (error) {
        console.error('Erro ao buscar colaboradores:', error);
        return [];
    }
}

// Salvar versão no histórico
async function saveVersion(firestoreId, action, spreadsheetData) {
    if (!currentUser || !firestoreId) return;

    try {
        // Converter array 2D para JSON string
        const versionData = {
            action: action,
            data: JSON.stringify(spreadsheetData),
            userId: currentUser.uid,
            userName: currentUser.displayName,
            timestamp: firebase.firestore.FieldValue.serverTimestamp()
        };

        await db.collection('spreadsheets').doc(firestoreId).collection('versions').add(versionData);
    } catch (error) {
        console.error('Erro ao salvar versão:', error);
    }
}

// Carregar histórico de versões
async function loadVersionHistory(firestoreId) {
    if (!currentUser || !firestoreId) return [];

    try {
        const snapshot = await db.collection('spreadsheets')
            .doc(firestoreId)
            .collection('versions')
            .orderBy('timestamp', 'desc')
            .get();

        const versions = [];
        snapshot.forEach((doc) => {
            const versionData = doc.data();
            versions.push({
                id: doc.id,
                ...versionData,
                // Converter JSON string de volta para objeto (se necessário para restaurar)
                // data: typeof versionData.data === 'string' ? JSON.parse(versionData.data) : versionData.data
            });
        });

        return versions;
    } catch (error) {
        console.error('Erro ao carregar versões:', error);
        return [];
    }
}
