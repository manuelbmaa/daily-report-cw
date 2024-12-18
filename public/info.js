// Credenciales predefinidas
const CREDENTIALS = {
    username: "admin",
    password: "M3g4c3ntr0.$", // Contraseña
};

// Crear una función para generar un hash simple (usamos btoa para codificar en base64)
function generateCredentialHash(username, password) {
    return btoa(username + ":" + password); // Codifica en base64
}

// Al enviar el formulario de login
document.getElementById("loginForm").addEventListener("submit", function (event) {
    event.preventDefault(); // Evitar recargar la página

    // Obtener los valores del formulario
    const username = document.getElementById("username").value.trim();
    const password = document.getElementById("password").value.trim();

    // Verificar si las credenciales coinciden con las predefinidas
    if (username === CREDENTIALS.username && password === CREDENTIALS.password) {
        // Generar el hash de las credenciales y almacenarlas en localStorage
        const currentHash = generateCredentialHash(CREDENTIALS.username, CREDENTIALS.password);
        localStorage.setItem("loggedIn", "true");
        localStorage.setItem("username", username); // Guardar el nombre de usuario
        localStorage.setItem("credentialHash", currentHash); // Guardar el hash de las credenciales

        // Redirigir a la página principal (o donde desees)
        window.location.href = "index.html"; // Asume que la página principal es "index.html"
    } else {
        // Si las credenciales no coinciden, mostrar el mensaje de error
        const errorMessage = document.getElementById("errorMessage");
        errorMessage.style.display = "block";
    }
});
