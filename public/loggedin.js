// Verificar si el usuario ha iniciado sesión
const storedHash = localStorage.getItem("credentialHash");
const currentHash = btoa("admin:1234"); // Actualiza las credenciales (usuario:contraseña) con las nuevas

if (!localStorage.getItem("loggedIn") || storedHash !== currentHash) {
    // Si no está autenticado o si las credenciales han cambiado, redirigir al login
    localStorage.removeItem("loggedIn");
    localStorage.removeItem("username");
    localStorage.removeItem("credentialHash");
    window.location.href = "login.html";
}
