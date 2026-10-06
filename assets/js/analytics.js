// Inicialización de analítica sin código inline, para permitir una CSP estricta.
window.dataLayer = window.dataLayer || [];
window.gtag = window.gtag || function () {
    window.dataLayer.push(arguments);
};
window.gtag('js', new Date());
window.gtag('config', 'G-RJ18T6PHS7');

(function (global, doc, key, scriptTag, projectId) {
    global[key] = global[key] || function () {
        (global[key].q = global[key].q || []).push(arguments);
    };
    const script = doc.createElement(scriptTag);
    script.async = true;
    script.src = 'https://www.clarity.ms/tag/' + projectId;
    const firstScript = doc.getElementsByTagName(scriptTag)[0];
    firstScript.parentNode.insertBefore(script, firstScript);
})(window, document, 'clarity', 'script', 'ymh5008s4u');
