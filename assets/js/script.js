
// =======================================================================
// CONFIGURACIÓN GLOBAL
// =======================================================================
const USE_LOCAL_XLSX = false; // true = usa el archivo local .xlsx | false = Google Sheets
const LOCAL_XLSX_FILE = 'BD_UAA EN NUM.xlsx';
// ID del documento de Google Sheets de donde se extrae la información
// Solo cambiar si el documento cambia
const SPREADSHEET_ID = '1ZWcV0JuWVLMn4NmZFcjGE6D0R16Wuar5477HOvgILew';

let localWorkbookPromise = null;

function getLocalWorkbook() {
    if (!localWorkbookPromise) {
    localWorkbookPromise = fetch(encodeURI(LOCAL_XLSX_FILE))
        .then(res => {
        if (!res.ok) throw new Error(`HTTP Error: ${res.status}`);
        return res.arrayBuffer();
        })
        .then(buffer => XLSX.read(buffer, { type: 'array' }));
    }
    return localWorkbookPromise;
}

// Función para obtener la URL de descarga del CSV
function getCsvUrl(sheetName, useExport = false) {
    if (useExport) {
    return `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/export?format=csv&sheet=${encodeURIComponent(sheetName)}`;
    }
    return `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(sheetName)}`;
}

function fetchSheetData(sheetName, useExport = false) {
    if (USE_LOCAL_XLSX) {
    return getLocalWorkbook().then(wb => {
        const sheet = wb.Sheets[sheetName];
        if (!sheet) {
        console.warn(`Hoja "${sheetName}" no encontrada en el XLSX.`);
        return [];
        }
        return XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: '',
        raw: false
        });
    });
    }

    return fetch(getCsvUrl(sheetName, useExport))
    .then(response => {
        if (!response.ok) throw new Error(`HTTP Error: ${response.status}`);
        return response.text();
    })
    .then(csvText => parseCSV(csvText));
}

function parseCSV(text) {
    let p = '', row = [''], ret = [row], i = 0, r = 0, s = !0, l;
    for (l of text) {
    if ('"' === l) {
        if (s && l === p) row[i] += l;
        s = !s;
    } else if (',' === l && s) l = row[++i] = '';
    else if ('\n' === l && s) {
        if ('\r' === p) row[i] = row[i].slice(0, -1);
        row = ret[++r] = [l = '']; i = 0;
    } else row[i] += l;
    p = l;
    }
    return ret;
}

function animarDato(elemento, objetivo, prefix = '') {
    if (!elemento) return;
    elemento.setAttribute('data-target', objetivo);
    const duration = 1400;
    const startTime = performance.now();
    const update = (currentTime) => {
    const progress = Math.min((currentTime - startTime) / duration, 1);
    const easeOut = 1 - Math.pow(1 - progress, 3);
    elemento.innerText = prefix + Math.floor(easeOut * objetivo).toLocaleString('es-MX');
    if (progress < 1) requestAnimationFrame(update);
    else elemento.innerText = prefix + objetivo.toLocaleString('es-MX');
    };
    requestAnimationFrame(update);
}

const cleanNum = (val) => {
    if (!val) return 0;
    const num = parseInt(String(val).replace(/[^0-9]/g, ''), 10);
    return isNaN(num) ? 0 : num;
};

// 1. Reveal Scroll
const reveals = document.querySelectorAll('.reveal');
const revealObserver = new IntersectionObserver((entries) => {
    entries.forEach(entry => {
    if (entry.isIntersecting) { entry.target.classList.add('active'); }
    });
}, { threshold: 0.1 });
reveals.forEach(el => revealObserver.observe(el));

// 2. Contadores Numéricos
const counters = document.querySelectorAll('.counter');
const countObserver = new IntersectionObserver((entries, observer) => {
    entries.forEach(entry => {
    if (entry.isIntersecting) {
        const counter = entry.target;
        const target = +counter.getAttribute('data-target');
        const prefix = counter.getAttribute('data-prefix') || '';
        const duration = 1400;
        const startTime = performance.now();
        const updateCount = (currentTime) => {
        const elapsed = currentTime - startTime;
        const progress = Math.min(elapsed / duration, 1);
        const easeOut = 1 - Math.pow(1 - progress, 3);
        
        counter.innerText = prefix + Math.floor(easeOut * target).toLocaleString('es-MX');
        
        if (progress < 1) { requestAnimationFrame(updateCount); } 
        else { counter.innerText = prefix + target.toLocaleString('es-MX'); }
        };
        requestAnimationFrame(updateCount);
        observer.unobserve(counter);
    }
    });
}, { threshold: 0.25 });
counters.forEach(counter => countObserver.observe(counter));

// 3. Activación de Barras, Dona y Pictograma
const graphicsObserver = new IntersectionObserver((entries) => {
    entries.forEach(entry => {
    if (entry.isIntersecting) {
        const target = entry.target;
        target.querySelectorAll('.progress-fill').forEach(bar => { 
        if(bar.hasAttribute('data-value') && bar.hasAttribute('data-total')) {
            bar.style.width = ((parseFloat(bar.getAttribute('data-value')) / parseFloat(bar.getAttribute('data-total'))) * 100) + '%';
        } else if (bar.hasAttribute('data-percent')) {
            bar.style.width = bar.getAttribute('data-percent'); 
        }
        });
        target.querySelectorAll('.donut-segment').forEach(seg => { seg.style.strokeDasharray = seg.getAttribute('data-dash'); });
        target.querySelectorAll('.bar-column, .history-bar-col').forEach(col => { col.style.height = col.getAttribute('data-height'); });
        
        if (target.classList.contains('card-gender-pictogram') && !target.dataset.animated) {
        target.dataset.animated = 'true';
        updateGenderPictogram('total');
        }
    }
    });
}, { threshold: 0.1 });
document.querySelectorAll('.card-bars, .card-donut, .offer-box, .card-12, .card-8, .card-7, .card-gender-pictogram').forEach(card => { graphicsObserver.observe(card); });

// 4. Acordeón Interactivo Seguro
const interactiveCards = document.querySelectorAll('.interactive-card:not(.card-hero-metric, .card-mini:has(.gender-details-wrapper))');
interactiveCards.forEach(card => {
    card.addEventListener('click', () => {
    interactiveCards.forEach(otherCard => {
        if (otherCard !== card) otherCard.classList.remove('active');
    });
    card.classList.toggle('active');
    });
});

// 5. Control de Modal
document.getElementById('modalCentros').addEventListener('click', function(e) {
    if (e.target === this) this.classList.remove('open');
});

// 6. Lógica del Pictograma Líquido
const genderData = {
    total: { men: 0, women: 0, x: 0, total: 0 },
    licenciatura: { men: 0, women: 0, x: 0, total: 0 },
    media: { men: 0, women: 0, x: 0, total: 0 },
    especialidades: { men: 0, women: 0, x: 0, total: 0 },
    maestria: { men: 0, women: 0, x: 0, total: 0 },
    doctorado: { men: 0, women: 0, x: 0, total: 0 }
};

function animateDecimal(id, targetValue) {
    const el = document.getElementById(id);
    const duration = 1200;
    let startTime = null;
    const update = (currentTime) => {
    if (!startTime) startTime = currentTime;
    const progress = Math.min((currentTime - startTime) / duration, 1);
    el.innerText = ((1 - Math.pow(1 - progress, 3)) * targetValue).toFixed(1) + '%';
    if (progress < 1) requestAnimationFrame(update);
    else el.innerText = targetValue.toFixed(1) + '%';
    };
    requestAnimationFrame(update);
}

function updateGenderPictogram(levelKey) {
    const data = genderData[levelKey] || genderData.total;
    const pMan = parseFloat(((data.men / data.total) * 100).toFixed(1));
    const pWoman = parseFloat(((data.women / data.total) * 100).toFixed(1));

    animateDecimal('manPercent', pMan);
    animateDecimal('womanPercent', pWoman);
    document.getElementById('genderCounts').innerText = `${data.men.toLocaleString('es-MX')} Hombres • ${data.women.toLocaleString('es-MX')} Mujeres`;
    
    const nonBinaryTag = document.getElementById('nonBinaryNotice');
    if (data.x > 0) {
    nonBinaryTag.style.display = 'inline-flex';
    document.getElementById('nonBinaryCount').innerText = data.x;
    } else {
    nonBinaryTag.style.display = 'none';
    }

    document.getElementById('manLiquid').setAttribute('y', 200 - ((pMan / 100) * 200));
    document.getElementById('manLiquid').setAttribute('height', (pMan / 100) * 200);
    document.getElementById('womanLiquid').setAttribute('y', 200 - ((pWoman / 100) * 200));
    document.getElementById('womanLiquid').setAttribute('height', (pWoman / 100) * 200);
}

const segmentButtons = document.querySelectorAll('#genderFilter .segment-btn');
segmentButtons.forEach(btn => {
    btn.addEventListener('click', () => {
    segmentButtons.forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    updateGenderPictogram(btn.getAttribute('data-level'));
    });
});

// 7. Lógica para Seleccionar Barras en Gráficas
function activarGraficaInteractiva(contenedorClase) {
    const container = document.querySelector(contenedorClase);
    if (!container) return;
    const groups = container.querySelectorAll('.chart-bar-group');
    groups.forEach(group => {
    group.addEventListener('click', () => {
        if (group.classList.contains('active')) {
        group.classList.remove('active');
        container.classList.remove('has-active');
        } else {
        groups.forEach(g => g.classList.remove('active'));
        group.classList.add('active');
        container.classList.add('has-active');
        }
    });
    });
}
activarGraficaInteractiva('.history-chart-container');
activarGraficaInteractiva('.bar-chart-container');

// --- NAVEGACIÓN FLOTANTE ARRASTRABLE (DRAG & DROP CON POINTER EVENTS) ---
(function initDraggableNav() {
const wrapper = document.getElementById('floatingNav');
const bubble = document.getElementById('navBubble');
const menu = document.getElementById('navMenu');
const links = menu.querySelectorAll('a');

let isDragging = false;
let hasMoved = false;
let startX = 0, startY = 0;
let initialLeft = 0, initialTop = 0;

bubble.addEventListener('pointerdown', (e) => {
isDragging = true;
hasMoved = false;
startX = e.clientX;
startY = e.clientY;

const rect = wrapper.getBoundingClientRect();
initialLeft = rect.left;
initialTop = rect.top;

// Fija la posición inline para liberar los atributos bottom/right de CSS
wrapper.style.bottom = 'auto';
wrapper.style.right = 'auto';
wrapper.style.left = `${initialLeft}px`;
wrapper.style.top = `${initialTop}px`;

bubble.setPointerCapture(e.pointerId);
});

bubble.addEventListener('pointermove', (e) => {
if (!isDragging) return;

const deltaX = e.clientX - startX;
const deltaY = e.clientY - startY;

// Umbral de 6px para diferenciar entre clic voluntario y arrastre accidental
if (Math.abs(deltaX) > 6 || Math.abs(deltaY) > 6) {
    hasMoved = true;
}

if (hasMoved) {
    // Límites de pantalla para que no se pierda fuera de la vista
    const maxX = window.innerWidth - wrapper.offsetWidth - 16;
    const maxY = window.innerHeight - wrapper.offsetHeight - 16;
    
    const newX = Math.min(Math.max(16, initialLeft + deltaX), maxX);
    const newY = Math.min(Math.max(16, initialTop + deltaY), maxY);

    wrapper.style.left = `${newX}px`;
    wrapper.style.top = `${newY}px`;

    // Ajusta la orientación del popover si la burbuja está muy arriba en la pantalla
    if (newY < 430) {
    menu.style.bottom = 'auto';
    menu.style.top = '70px';
    menu.style.transformOrigin = 'top right';
    } else {
    menu.style.top = 'auto';
    menu.style.bottom = '70px';
    menu.style.transformOrigin = 'bottom right';
    }
}
});

const endDrag = (e) => {
if (!isDragging) return;
isDragging = false;
bubble.releasePointerCapture(e.pointerId);

// Si no hubo desplazamiento, fue un clic: alternar menú
if (!hasMoved) {
    const isOpen = menu.classList.toggle('open');
    wrapper.classList.toggle('menu-open', isOpen);
}
};

bubble.addEventListener('pointerup', endDrag);
bubble.addEventListener('pointercancel', endDrag);

// Cierre y scroll suave al pulsar una opción del menú
links.forEach(link => {
link.addEventListener('click', (e) => {
    e.preventDefault();
    const targetId = link.getAttribute('href');
    const targetSection = document.querySelector(targetId);

    menu.classList.remove('open');
    wrapper.classList.remove('menu-open');

    if (targetSection) {
        targetSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
        history.pushState(null, '', targetId);
    }
});
});

// Cerrar menú al hacer clic fuera
document.addEventListener('pointerdown', (e) => {
if (!wrapper.contains(e.target) && menu.classList.contains('open')) {
    menu.classList.remove('open');
    wrapper.classList.remove('menu-open');
}
});
})();

// --- SECUENCIA DE ONBOARDING DEL TOOLTIP (SIN AUTO-ABRIR MENÚ) ---
setTimeout(() => {
const wrapper = document.getElementById('floatingNav');
if (!wrapper) return;

// Mostrar solo el tooltip indicativo por 4 segundos
wrapper.classList.add('show-tooltip');
setTimeout(() => {
    wrapper.classList.remove('show-tooltip');
}, 4000);
}, 2000);

// --- RECORDATORIO PERIÓDICO DEL TOOLTIP ---
setInterval(() => {
const wrapper = document.getElementById('floatingNav');
if (wrapper && !wrapper.classList.contains('menu-open')) {
    wrapper.classList.add('show-tooltip');
    setTimeout(() => {
        wrapper.classList.remove('show-tooltip');
    }, 4000);
}
}, 25000);

// =========================================================================
// SISTEMA DE ANCLAS CON SCROLL SMOOTH Y COPIADO DE ENLACE EN HEADER
// =========================================================================
let copyToastTimeout;
function showCopyToast(msg) {
    const toast = document.getElementById('copyToast');
    const text = document.getElementById('copyToastText');
    if (!toast) return;
    if (text) text.textContent = msg;
    toast.classList.add('show');
    clearTimeout(copyToastTimeout);
    copyToastTimeout = setTimeout(() => {
        toast.classList.remove('show');
    }, 2800);
}

function copyTextToClipboard(text) {
    if (navigator.clipboard && window.isSecureContext) {
        return navigator.clipboard.writeText(text);
    } else {
        const textArea = document.createElement('textarea');
        textArea.value = text;
        textArea.style.position = 'fixed';
        textArea.style.left = '-999999px';
        textArea.style.top = '-999999px';
        document.body.appendChild(textArea);
        textArea.focus();
        textArea.select();
        return new Promise((resolve, reject) => {
            const ok = document.execCommand('copy');
            textArea.remove();
            ok ? resolve() : reject();
        });
    }
}

function initSmoothScrollAndHeaderLinks() {
    // 1. Delegación para cualquier enlace interno a anclas (#)
    document.addEventListener('click', (e) => {
        const link = e.target.closest('a[href^="#"]');
        if (!link) return;

        const targetId = link.getAttribute('href');
        if (!targetId || targetId === '#') return;

        const targetEl = document.querySelector(targetId);
        if (targetEl) {
            e.preventDefault();
            targetEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
            history.pushState(null, '', targetId);
        }
    });

    // 2. Soporte para URL con hash al cargar la página
    if (window.location.hash) {
        const scrollToHash = () => {
            const targetEl = document.querySelector(window.location.hash);
            if (targetEl) {
                targetEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
            }
        };
        if (document.readyState === 'complete') {
            setTimeout(scrollToHash, 300);
        } else {
            window.addEventListener('load', () => setTimeout(scrollToHash, 300));
        }
    }

    // 3. Soporte para evento hashchange
    window.addEventListener('hashchange', () => {
        if (window.location.hash) {
            const targetEl = document.querySelector(window.location.hash);
            if (targetEl) {
                targetEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
            }
        }
    });

    // 4. Inyección del botón de copiado en cada header#bloque-
    const headers = document.querySelectorAll('header.block-header[id]');
    headers.forEach(header => {
        const tag = header.querySelector('.block-tag');
        if (!tag || tag.querySelector('.block-anchor-btn')) return;

        const copyBtn = document.createElement('button');
        copyBtn.className = 'block-anchor-btn';
        copyBtn.type = 'button';
        copyBtn.setAttribute('aria-label', `Copiar enlace a ${header.id}`);
        copyBtn.title = 'Copiar enlace a este bloque';
        copyBtn.innerHTML = `
            <svg class="anchor-icon-link" viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="2.3" stroke-linecap="round" stroke-linejoin="round">
                <path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"></path>
                <path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"></path>
            </svg>
            <svg class="anchor-icon-check" viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" style="display: none;">
                <polyline points="20 6 9 17 4 12"></polyline>
            </svg>
        `;

        tag.appendChild(copyBtn);

        const handleCopy = (e) => {
            e.preventDefault();
            e.stopPropagation();
            const blockId = header.id;
            // Extraer el texto del tag ignorando el botón de ancla
            const tagSpan = tag.querySelector('span:nth-of-type(2)');
            const title = tagSpan 
                ? tagSpan.textContent.trim() 
                : (tag.childNodes[1]?.textContent?.trim() || header.querySelector('.block-title')?.textContent?.trim() || blockId);
            const fullUrl = `${window.location.origin}${window.location.pathname}${window.location.search}#${blockId}`;

            copyTextToClipboard(fullUrl).then(() => {
                history.pushState(null, '', `#${blockId}`);
                
                copyBtn.classList.add('copied');
                const linkIcon = copyBtn.querySelector('.anchor-icon-link');
                const checkIcon = copyBtn.querySelector('.anchor-icon-check');
                if (linkIcon && checkIcon) {
                    linkIcon.style.display = 'none';
                    checkIcon.style.display = 'block';
                }

                showCopyToast(`Enlace copiado: ${title}`);

                setTimeout(() => {
                    copyBtn.classList.remove('copied');
                    if (linkIcon && checkIcon) {
                        linkIcon.style.display = 'block';
                        checkIcon.style.display = 'none';
                    }
                }, 2000);
            }).catch(err => {
                console.error('Error al copiar enlace:', err);
            });
        };

        copyBtn.addEventListener('click', handleCopy);
        tag.addEventListener('click', handleCopy);
    });
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initSmoothScrollAndHeaderLinks);
} else {
    initSmoothScrollAndHeaderLinks();
}

// =========================================================================
// AUTOMATIZACIÓN BLOQUE 01: MATRÍCULA
// =========================================================================
document.addEventListener('DOMContentLoaded', () => {
fetchSheetData('01_Matricula', true).then(data => {
    
    let totalGlobal = 0, masGlobal = 0, femGlobal = 0, xGlobal = 0;
    let textosInterfase = {};
    let centrosAcademicos = [];
    
    const nivelesResumen = {
        'Educación Media': { mas: 0, fem: 0, x: 0, total: 0 },
        'Licenciatura': { mas: 0, fem: 0, x: 0, total: 0 },
        'Especialidades': { mas: 0, fem: 0, x: 0, total: 0 },
        'Especialidades Médicas': { mas: 0, fem: 0, x: 0, total: 0 },
        'Maestría': { mas: 0, fem: 0, x: 0, total: 0 },
        'Doctorado': { mas: 0, fem: 0, x: 0, total: 0 }
    };

    // CORRECCIÓN: Ahora el ciclo empieza en 0 para no saltarnos la fila 1
    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        if (!row || row.length === 0) continue;

        // Mapeo Textos Naranjas (Col G -> row[6])
        const keyTexto = (row[6] || '').trim();
        if (keyTexto) {
        textosInterfase[keyTexto] = { 
            h: (row[7] || '').trim(), 
            i: (row[8] || '').trim(), 
            j: (row[9] || '').trim() 
        };
        }

        // Mapeo Números (Col A-E -> row[0] a row[4])
        const nivelOriginal = (row[0] || '').trim();
        
        // CORRECCIÓN: Le decimos que ignore específicamente la celda del encabezado para no sumarla
        if (nivelOriginal && nivelOriginal !== 'Nivel Académico') {
        const centro = (row[1] || '').trim();
        const mas = cleanNum(row[2]);
        const fem = cleanNum(row[3]);
        const x = cleanNum(row[4]);
        const filaTotal = mas + fem + x;

        totalGlobal += filaTotal; 
        masGlobal += mas; 
        femGlobal += fem; 
        xGlobal += x;

        let nivelClave = nivelOriginal;
        if (nivelesResumen[nivelClave]) {
            nivelesResumen[nivelClave].mas += mas;
            nivelesResumen[nivelClave].fem += fem;
            nivelesResumen[nivelClave].x += x;
            nivelesResumen[nivelClave].total += filaTotal;
        }
        centrosAcademicos.push({ nivelOriginal, centro, mas, fem, x, filaTotal });
        }
    }

    // 1. Textos
    const b1 = document.getElementById('bloque-01');
    if (b1) {
        if (textosInterfase['PRIMERA LÍNEA']) {
            const b1Tag = document.getElementById('b1-tag-text') || b1.querySelector('.block-tag');
            if (b1Tag) b1Tag.innerText = textosInterfase['PRIMERA LÍNEA'].h;
        }
        if (textosInterfase['LÍNEA PRINCIPAL']) b1.querySelector('.block-title').innerText = textosInterfase['LÍNEA PRINCIPAL'].h;
        if (textosInterfase['INTRODUCCIÓN']) {
        b1.querySelector('.block-lead').innerHTML = `${textosInterfase['INTRODUCCIÓN'].h} <strong>${totalGlobal.toLocaleString('es-MX')} estudiantes</strong> ${textosInterfase['INTRODUCCIÓN'].i} <strong>${textosInterfase['INTRODUCCIÓN'].j}</strong>.`;
        }
        // NUEVO: Fecha de corte automatizada
        if (textosInterfase['FECHA DE CORTE']) {
        b1.querySelector('.block-date').innerText = `Corte al ${textosInterfase['FECHA DE CORTE'].h.toLowerCase()}`;
        }
    }

    const heroLbl = document.querySelector('.hero-metric-label');
    const heroDesc = document.querySelector('.hero-metric-desc');
    const genderTitle = document.querySelector('.gender-card-title-group h3');
    const genderSub = document.querySelector('.gender-card-title-group p');

    if (heroLbl && textosInterfase['TARJETA 1 PRIMERA LÍNEA']) heroLbl.innerHTML = textosInterfase['TARJETA 1 PRIMERA LÍNEA'].h;
    if (heroDesc && textosInterfase['TARJETA 1 PRINCIPAL']) heroDesc.innerText = textosInterfase['TARJETA 1 PRINCIPAL'].h;
    if (genderTitle && textosInterfase['TÍTULO GRÁFICO']) genderTitle.innerText = textosInterfase['TÍTULO GRÁFICO'].h;
    if (genderSub && textosInterfase['SUBTÍTULO']) genderSub.innerText = textosInterfase['SUBTÍTULO'].h;

    // 2. Métrica Central
    animarDato(document.querySelector('.hero-metric-num'), totalGlobal);

    // 3. Tarjetas Mini
    const miniCards = document.querySelectorAll('.card-mini');
    const ordenTarjetas = ['Educación Media', 'Licenciatura', 'Especialidades', 'Especialidades Médicas', 'Maestría', 'Doctorado'];
    miniCards.forEach((card, index) => {
        const datosNivel = nivelesResumen[ordenTarjetas[index]];
        if (!datosNivel) return;
        animarDato(card.querySelector('.mini-num'), datosNivel.total);
        const generosVal = card.querySelectorAll('.gender-val');
        if (generosVal.length >= 3) {
        generosVal[0].innerText = datosNivel.mas.toLocaleString('es-MX');
        generosVal[1].innerText = datosNivel.fem.toLocaleString('es-MX');
        generosVal[2].innerText = datosNivel.x.toLocaleString('es-MX');
        }
    });

    // 4. Género
    if (typeof genderData !== 'undefined') {
        genderData.total = { men: masGlobal, women: femGlobal, x: xGlobal, total: totalGlobal };
        genderData.licenciatura = { men: nivelesResumen['Licenciatura'].mas, women: nivelesResumen['Licenciatura'].fem, x: nivelesResumen['Licenciatura'].x, total: nivelesResumen['Licenciatura'].total };
        genderData.media = { men: nivelesResumen['Educación Media'].mas, women: nivelesResumen['Educación Media'].fem, x: nivelesResumen['Educación Media'].x, total: nivelesResumen['Educación Media'].total };
        genderData.especialidades = { men: nivelesResumen['Especialidades'].mas, women: nivelesResumen['Especialidades'].fem, x: nivelesResumen['Especialidades'].x, total: nivelesResumen['Especialidades'].total };
        genderData.especialidadesMedicas = { men: nivelesResumen['Especialidades Médicas'].mas, women: nivelesResumen['Especialidades Médicas'].fem, x: nivelesResumen['Especialidades Médicas'].x, total: nivelesResumen['Especialidades Médicas'].total };
        genderData.maestria = { men: nivelesResumen['Maestría'].mas, women: nivelesResumen['Maestría'].fem, x: nivelesResumen['Maestría'].x, total: nivelesResumen['Maestría'].total };
        genderData.doctorado = { men: nivelesResumen['Doctorado'].mas, women: nivelesResumen['Doctorado'].fem, x: nivelesResumen['Doctorado'].x, total: nivelesResumen['Doctorado'].total };

        const activeBtn = document.querySelector('#genderFilter .segment-btn.active');
        if (activeBtn && typeof updateGenderPictogram === 'function') {
        updateGenderPictogram(activeBtn.getAttribute('data-level'));
        }
    }

    // 5. Tabla Modal
    let subtotalesModal = {};
    centrosAcademicos.forEach(c => {
        if (!subtotalesModal[c.nivelOriginal]) subtotalesModal[c.nivelOriginal] = { mas: 0, fem: 0, x: 0, total: 0 };
        subtotalesModal[c.nivelOriginal].mas += c.mas;
        subtotalesModal[c.nivelOriginal].fem += c.fem;
        subtotalesModal[c.nivelOriginal].x += c.x;
        subtotalesModal[c.nivelOriginal].total += c.filaTotal;
    });

    let htmlTabla = '';
    let nivelActual = '';
    centrosAcademicos.forEach(c => {
        if (c.nivelOriginal !== nivelActual) {
        nivelActual = c.nivelOriginal;
        const sub = subtotalesModal[nivelActual];
        htmlTabla += `<tr class="row-level"><td>${nivelActual.toUpperCase()}</td><td class="text-right">${sub.total.toLocaleString('es-MX')}</td><td class="text-right col-gender">${sub.mas.toLocaleString('es-MX')}</td><td class="text-right col-gender">${sub.fem.toLocaleString('es-MX')}</td><td class="text-right col-gender">${sub.x.toLocaleString('es-MX')}</td></tr>`;
        }
        htmlTabla += `<tr class="row-center"><td>${c.centro}</td><td class="text-right">${c.filaTotal.toLocaleString('es-MX')}</td><td class="text-right col-gender">${c.mas.toLocaleString('es-MX')}</td><td class="text-right col-gender">${c.fem.toLocaleString('es-MX')}</td><td class="text-right col-gender">${c.x.toLocaleString('es-MX')}</td></tr>`;
    });
    
    const tbody = document.querySelector('.glass-table tbody');
    if (tbody) tbody.innerHTML = htmlTabla;
    })
    .catch(error => {
    console.error('Detalle del error en Bloque 01:', error);
    });
});

// =========================================================================
// AUTOMATIZACIÓN BLOQUE 02: OFERTA EDUCATIVA
// =========================================================================
document.addEventListener('DOMContentLoaded', () => {
fetchSheetData('02_Oferta_Educativa').then(data => {
    
    let textosInterfase = {};
    let datosNiveles = [];
    
    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        if (!row || row.length === 0) continue;

        // Mapeo Textos Naranjas (Col G, H, I, J -> Índices 6, 7, 8, 9)
        const keyTexto = (row[6] || '').trim();
        if (keyTexto) {
        textosInterfase[keyTexto] = { 
            h: (row[7] || '').trim(), 
            i: (row[8] || '').trim(), 
            j: (row[9] || '').trim() 
        };
        }

        // Mapeo Números (Col A, B, C -> Índices 0, 1, 2)
        const nivel = (row[0] || '').trim();
        if (nivel && nivel !== 'Nivel') {
        const prog = cleanNum(row[1]);
        const lug = cleanNum(row[2]);
        datosNiveles.push({ nivel, prog, lug });
        }
    }

    // ================= ACTUALIZACIÓN DEL HTML =================
    const b2 = document.getElementById('bloque-02');
    if (!b2) return;

    // 1. Textos Generales
    if (textosInterfase['PRIMERA LÍNEA']) {
        const b2Tag = document.getElementById('b2-tag-text') || b2.querySelector('.block-tag');
        if (b2Tag) b2Tag.innerText = textosInterfase['PRIMERA LÍNEA'].h;
    }
    if (textosInterfase['LÍNEA PRINCIPAL']) b2.querySelector('.block-title').innerText = textosInterfase['LÍNEA PRINCIPAL'].h;
    if (textosInterfase['INTRODUCCIÓN']) {
        b2.querySelector('.block-lead').innerHTML = `${textosInterfase['INTRODUCCIÓN'].h} <strong>${textosInterfase['INTRODUCCIÓN'].i}</strong>, ${textosInterfase['INTRODUCCIÓN'].j}`;
    }

    // 2. Links / Botones de acción
    const links = b2.querySelectorAll('.block-action-link');
    if (textosInterfase['LINK 1'] && links.length > 0) {
        links[0].innerHTML = `${textosInterfase['LINK 1'].h} &#8599;`;
        links[0].href = textosInterfase['LINK 1'].i;
    }
    if (textosInterfase['LINK 2'] && links.length > 1) {
        links[1].innerHTML = `${textosInterfase['LINK 2'].h} &#8599;`;
        links[1].href = textosInterfase['LINK 2'].i;
    }

    // 3. Programas y Lugares Ofertados (Usando document para abarcar toda la sección)
    const numProgramas = document.querySelectorAll('.offer-box .program-num');
    const numLugares = document.querySelectorAll('.bar-chart-container .bar-value-label');
    const barras = document.querySelectorAll('.bar-chart-container .bar-column');

    const maxLugares = Math.max(...datosNiveles.map(d => d.lug));

    datosNiveles.forEach((dato, index) => {
        if (numProgramas[index]) animarDato(numProgramas[index], dato.prog);
        if (numLugares[index]) animarDato(numLugares[index], dato.lug);
        
        if (barras[index]) {
        const porcentaje = maxLugares > 0 ? (dato.lug / maxLugares) * 100 : 0;
        barras[index].setAttribute('data-height', porcentaje + '%');
        barras[index].style.height = porcentaje + '%';
        }
    });

    })
    .catch(error => {
    console.error('Detalle del error en Bloque 02:', error);
    });
});

// =========================================================================
// AUTOMATIZACIÓN BLOQUE 03: CALIDAD Y RECONOCIMIENTO
// =========================================================================
document.addEventListener('DOMContentLoaded', () => {
fetchSheetData('03_Acreditaciones').then(data => {
    
    let textosInterfase = {};
    let filasAcreditaciones = [];

    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        if (!row || row.length === 0) continue;

        const keyTexto = (row[6] || '').trim();
        if (keyTexto) {
        const textoCompleto = row.slice(7).filter(t => t && t.trim() !== '').join(' ');
        textosInterfase[keyTexto] = textoCompleto;
        }

        const nivel = (row[0] || '').trim();
        if (nivel && nivel !== 'Nivel Académico') {
        filasAcreditaciones.push({
            nivel,
            tipo: (row[1] || '').trim(),
            total: cleanNum(row[2]),
            ofertables: cleanNum(row[3]),
            acreditados: cleanNum(row[4])
        });
        }
    }

    // ================= ACTUALIZACIÓN DEL HTML =================
    const b3 = document.getElementById('bloque-03');
    if (!b3) return;

    // 1. Textos Generales del Bloque
    if (textosInterfase['PRIMERA LÍNEA']) {
        const b3Tag = document.getElementById('b3-tag-text') || b3.querySelector('.block-tag');
        if (b3Tag) b3Tag.innerText = textosInterfase['PRIMERA LÍNEA'];
    }
    if (textosInterfase['LÍNEA PRINCIPAL']) b3.querySelector('.block-title').innerText = textosInterfase['LÍNEA PRINCIPAL'];
    if (textosInterfase['INTRODUCCIÓN']) b3.querySelector('.block-lead').innerHTML = textosInterfase['INTRODUCCIÓN'];
    if (textosInterfase['FECHA DE CORTE']) b3.querySelector('.block-date').innerText = `Corte al ${textosInterfase['FECHA DE CORTE'].toLowerCase()}`;

    // 2. Tarjeta 1 (Reconstruyendo el texto con los números 64, 63 y el 100%)
    if (filasAcreditaciones.length > 0) {
        const t1 = filasAcreditaciones[0];
        animarDato(document.querySelector('.t1-num'), t1.acreditados);
        
        const t1Total = document.querySelector('.t1-total');
        if (t1Total) t1Total.innerText = t1.ofertables;
        
        animarDato(document.querySelector('.t1-pct'), 100);

        // Inyectar párrafo dinámico con el total (64) y acreditados (63)
        const t1Desc = document.querySelector('.t1-desc');
        if (t1Desc) {
        t1Desc.innerHTML = `Del total de <strong>${t1.total} programas educativos</strong> de licenciatura, ${t1.ofertables} son evaluables; el <strong>100%</strong> de ellos están acreditados y reconocidos por algún organismo nacional.`;
        }
    }
    if (textosInterfase['TARJETA 1 PRIMERA LÍNEA']) document.querySelector('.t1-tag').innerText = textosInterfase['TARJETA 1 PRIMERA LÍNEA'];
    if (textosInterfase['TARJETA 1 PRINCIPAL']) document.querySelector('.t1-title').innerText = textosInterfase['TARJETA 1 PRINCIPAL'];
    if (textosInterfase['TARJETA 1 SECUNDARIA']) document.querySelector('.t1-sub').innerText = textosInterfase['TARJETA 1 SECUNDARIA'];

    // 3. Tarjeta 2 (CONAHCYT - 100%)
    if (textosInterfase['TARJETA 2 PRIMERA LÍNEA']) document.querySelector('.t2-tag').innerText = textosInterfase['TARJETA 2 PRIMERA LÍNEA'];
    if (textosInterfase['TARJETA 2 PRINCIPAL']) document.querySelector('.t2-title').innerText = textosInterfase['TARJETA 2 PRINCIPAL'];
    
    // Calcular el porcentaje real de posgrados (FilasAcreditaciones[1] corresponde a Posgrado/CONAHCYT)
    let porcentajeT2 = 100; // Valor por defecto
    if (filasAcreditaciones.length > 1) {
        const t2 = filasAcreditaciones[1];
        porcentajeT2 = t2.ofertables > 0 ? Math.round((t2.acreditados / t2.ofertables) * 100) : 0;
    }

    // Reconstruir el texto inyectando el porcentaje y el nombre del sistema en negritas
    const t2Desc = document.querySelector('.t2-desc');
    if (t2Desc) {
        t2Desc.innerHTML = `El <strong>${porcentajeT2}%</strong> de nuestros programas educativos de posgrado ofertables se encuentran avalados y vigentes en el <strong>Sistema Nacional de Posgrados (SNP)</strong>.`;
    }
    if (textosInterfase['TARJETA 2 COMENTARIO']) document.querySelector('.t2-footer').innerText = textosInterfase['TARJETA 2 COMENTARIO'];
    
    // Animamos el gráfico con la variable calculada en lugar del 100 fijo
    animarDato(document.querySelector('.t2-pct'), porcentajeT2);

    // 4. Tarjeta 3 (Internacional)
    if (filasAcreditaciones.length > 2) {
        animarDato(document.querySelector('.t3-num'), filasAcreditaciones[2].acreditados);
    }
    if (textosInterfase['TARJETA 3 PRIMERA LÍNEA']) document.querySelector('.t3-tag').innerText = textosInterfase['TARJETA 3 PRIMERA LÍNEA'];
    if (textosInterfase['TARJETA 3 PRINCIPAL']) document.querySelector('.t3-title').innerText = textosInterfase['TARJETA 3 PRINCIPAL'];
    if (textosInterfase['TARJETA 3 SECUNDARIA']) document.querySelector('.t3-sub').innerText = textosInterfase['TARJETA 3 SECUNDARIA'];
    if (textosInterfase['TARJETA 3 COMENTARIO']) document.querySelector('.t3-desc').innerHTML = textosInterfase['TARJETA 3 COMENTARIO'];

    // 5. Tarjeta 4 (Padrón EGEL)
    if (filasAcreditaciones.length > 3) {
        animarDato(document.querySelector('.t4-num'), filasAcreditaciones[3].acreditados);
    }
    if (textosInterfase['TARJETA 4 PRIMERA LÍNEA']) document.querySelector('.t4-tag').innerText = textosInterfase['TARJETA 4 PRIMERA LÍNEA'];
    if (textosInterfase['TARJETA 4 PRINCIPAL']) document.querySelector('.t4-title').innerText = textosInterfase['TARJETA 4 PRINCIPAL'];
    if (textosInterfase['TARJETA 4 SECUNDARIA']) document.querySelector('.t4-sub').innerText = textosInterfase['TARJETA 4 SECUNDARIA'];
    if (textosInterfase['TARJETA 4 COMENTARIO']) document.querySelector('.t4-desc').innerHTML = textosInterfase['TARJETA 4 COMENTARIO'];

    // 6. Tarjeta 5 (Posgrados Globales)
    if (filasAcreditaciones.length > 4) {
        animarDato(document.querySelector('.t5-num'), filasAcreditaciones[4].acreditados);
    }
    if (textosInterfase['TARJETA 5 PRIMERA LÍNEA']) document.querySelector('.t5-tag').innerText = textosInterfase['TARJETA 5 PRIMERA LÍNEA'];
    if (textosInterfase['TARJETA 5 PRINCIPAL']) document.querySelector('.t5-title').innerText = textosInterfase['TARJETA 5 PRINCIPAL'];
    if (textosInterfase['TARJETA 5 SECUNDARIA']) document.querySelector('.t5-sub').innerText = textosInterfase['TARJETA 5 SECUNDARIA'];
    if (textosInterfase['TARJETA 5 COMENTARIO']) document.querySelector('.t5-desc').innerHTML = textosInterfase['TARJETA 5 COMENTARIO'];

    })
    .catch(error => {
    console.error('Detalle del error en Bloque 03:', error);
    });
});

// =========================================================================
// AUTOMATIZACIÓN BLOQUE 04: BECAS Y APOYOS (GRÁFICA HISTÓRICA DINÁMICA)
// =========================================================================
window.addEventListener('DOMContentLoaded', () => {
fetchSheetData('04_Becas').then(data => {
    
    let textosInterfase = {};
    let inversiones = [];
    let beneficiarios = [];
    let historico = [];

    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        if (!row || row.length === 0) continue;

        const keyTexto = (row[6] || '').trim();
        if (keyTexto) {
        textosInterfase[keyTexto] = { h: (row[7] || '').trim(), i: (row[8] || '').trim() };
        }

        let cat = (row[0] || '').trim();
        let rubroCol = (row[1] || '').trim();
        let valorCol = cleanNum(row[2]);

        if (cat && cat !== 'Categoría') {
        let catLower = cat.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
        if (catLower.includes('inversio')) {
            inversiones.push({ rubro: rubroCol, valor: valorCol });
        } else if (catLower.includes('beneficiar')) {
            beneficiarios.push({ rubro: rubroCol, valor: valorCol });
        } else if (catLower.includes('historico')) {
            historico.push({ anio: rubroCol, valor: valorCol });
        }
        }
    }

    // 1. Textos Generales
    if (textosInterfase['PRIMERA LÍNEA']) document.getElementById('b4-tag-text').innerText = textosInterfase['PRIMERA LÍNEA'].h;
    if (textosInterfase['LÍNEA PRINCIPAL']) document.getElementById('b4-title').innerText = textosInterfase['LÍNEA PRINCIPAL'].h;
    if (textosInterfase['INTRODUCCIÓN']) document.getElementById('b4-lead').innerText = textosInterfase['INTRODUCCIÓN'].h;
    if (textosInterfase['FECHA DE CORTE']) document.getElementById('b4-date').innerText = `Corte al ${textosInterfase['FECHA DE CORTE'].h.toLowerCase()}`;
    if (textosInterfase['LINK']) {
        const link = document.getElementById('b4-link');
        link.innerHTML = `${textosInterfase['LINK'].h} &#8599;`;
        link.href = textosInterfase['LINK'].i;
    }

    // 2. Tarjeta Inversión Total
    if (textosInterfase['TARJETA 1 PRIMERA LÍNEA']) document.getElementById('b4-t1-tag').innerText = textosInterfase['TARJETA 1 PRIMERA LÍNEA'].h;
    let totalInversion = inversiones.reduce((sum, item) => sum + item.valor, 0);
    document.getElementById('b4-total-inv').setAttribute('data-target', totalInversion);

    let htmlBreakdown = '';
    inversiones.forEach(inv => {
        htmlBreakdown += `
        <div class="breakdown-item">
            <div class="breakdown-header"><span class="breakdown-label">${inv.rubro}</span><span class="breakdown-val counter" data-target="${inv.valor}" data-prefix="$">0</span></div>
            <div class="progress-track"><div class="progress-fill" data-value="${inv.valor}" data-total="${totalInversion}" style="background: #565D6D; width: 0%;"></div></div>
        </div>`;
    });
    document.getElementById('b4-breakdown').innerHTML = htmlBreakdown;

    // 3. Tarjeta Beneficiarios
    if (textosInterfase['TARJETA 2 PRIMERA LÍNEA']) document.getElementById('b4-t2-tag').innerText = textosInterfase['TARJETA 2 PRIMERA LÍNEA'].h;
    if (textosInterfase['TARJETA 2 SECUNDARIA']) document.getElementById('b4-t2-desc').innerText = textosInterfase['TARJETA 2 SECUNDARIA'].h;

    let totalBeneficiarios = beneficiarios.reduce((sum, item) => sum + item.valor, 0);
    document.getElementById('b4-total-beneficiarios').setAttribute('data-target', totalBeneficiarios);

    let htmlBeneficiarios = '';
    beneficiarios.forEach(ben => {
        htmlBeneficiarios += `
        <div class="beneficiary-row"><span class="beneficiary-num counter" data-target="${ben.valor}">0</span><span class="beneficiary-text">${ben.rubro}</span></div>`;
    });
    document.getElementById('b4-beneficiarios').innerHTML = htmlBeneficiarios;

    // 4. Histórico de Becas
    const minYear = historico.length > 0 ? historico[0].anio : '2018';
    const maxYear = historico.length > 0 ? historico[historico.length - 1].anio : '2025';
    document.getElementById('b4-t3-tag').innerText = `Histórico de Becas Otorgadas (${minYear} - ${maxYear})`;

    const maxHistorico = Math.max(...historico.map(h => h.valor), 0);
    let htmlBars = '';
    let htmlXAxis = '';

    historico.forEach(hist => {
        const porcentaje = maxHistorico > 0 ? (hist.valor / maxHistorico) * 100 : 0;
        htmlBars += `
        <div class="chart-bar-group"><span class="history-val-label counter" data-target="${hist.valor}">0</span><div class="history-bar-col" data-height="${porcentaje.toFixed(1)}%" style="height: 0%;"></div></div>`;
        htmlXAxis += `<div class="x-axis-label">${hist.anio}</div>`;
    });

    const historyChart = document.getElementById('b4-history-chart');
    const historyXaxis = document.getElementById('b4-history-xaxis');

    historyChart.innerHTML = htmlBars;
    historyXaxis.innerHTML = htmlXAxis;

    // ESTA ES LA SOLUCIÓN: Sobrescribe el CSS para que la cuadrícula 
    // tenga exactamente tantas columnas como años haya en Google Sheets
    historyChart.style.gridTemplateColumns = `repeat(${historico.length}, 1fr)`;
    historyXaxis.style.gridTemplateColumns = `repeat(${historico.length}, 1fr)`;

    // =========================================================================
    // DEVOLVER ANIMACIONES E INTERACTIVIDAD AL DOM INYECTADO
    // =========================================================================
    if (historyChart) {
        const groups = historyChart.querySelectorAll('.chart-bar-group');
        groups.forEach(group => {
        group.addEventListener('click', () => {
            if (group.classList.contains('active')) {
            group.classList.remove('active');
            historyChart.classList.remove('has-active');
            } else {
            groups.forEach(g => g.classList.remove('active'));
            group.classList.add('active');
            historyChart.classList.add('has-active');
            }
        });
        });
    }

    setTimeout(() => {
        const cardsContainer = document.getElementById('bloque-04-cards');
        if(!cardsContainer) return;

        cardsContainer.querySelectorAll('.counter').forEach(c => {
        const target = parseFloat(c.getAttribute('data-target') || 0);
        const prefix = c.getAttribute('data-prefix') || '';
        animarDato(c, target, prefix);
        });

        cardsContainer.querySelectorAll('.progress-fill').forEach(f => {
        const val = parseFloat(f.getAttribute('data-value') || 0);
        const tot = parseFloat(f.getAttribute('data-total') || 1);
        if(tot > 0) f.style.width = (val / tot * 100) + '%';
        });

        cardsContainer.querySelectorAll('.history-bar-col').forEach(b => {
        b.style.height = b.getAttribute('data-height');
        });
    }, 50); 

    })
    .catch(error => {
    console.error('Detalle del error en Bloque 04:', error);
    });
});

// =========================================================================
// AUTOMATIZACIÓN BLOQUE 05: EGRESADOS E IMPACTO HISTÓRICO (CORREGIDO)
// =========================================================================
window.addEventListener('DOMContentLoaded', () => {
fetchSheetData('05_Egresados').then(data => {
    
    let textosInterfase = {};
    let nivelesEgresados = [];

    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        if (!row || row.length === 0) continue;

        // Textos de configuración (Col G, H e I)
        const keyTexto = (row[6] || '').trim();
        if (keyTexto) {
        textosInterfase[keyTexto] = { 
            h: (row[7] || '').trim(),
            i: (row[8] || '').trim() // Capturamos la columna I
        };
        }

        // Datos de la tabla (Col A y B/C)
        let nivelTexto = (row[0] || '').trim();
        let egresadosVal = cleanNum(row[1]) || cleanNum(row[2]); 

        if (nivelTexto && !nivelTexto.toLowerCase().includes('nivel') && egresadosVal > 0) {
        nivelesEgresados.push({ nombre: nivelTexto, valor: egresadosVal });
        }
    }

    const b5 = document.getElementById('bloque-05-cards');
    if (!b5) return;

    // 1. Textos Generales
    if (textosInterfase['PRIMERA LÍNEA']) document.getElementById('b5-tag-text').innerText = textosInterfase['PRIMERA LÍNEA'].h;
    if (textosInterfase['LÍNEA PRINCIPAL']) document.getElementById('b5-title').innerText = textosInterfase['LÍNEA PRINCIPAL'].h;
    
    // CORRECCIÓN: Armamos la introducción combinando Col H y Col I (con etiqueta strong)
    if (textosInterfase['INTRODUCCIÓN']) {
        let introHTML = textosInterfase['INTRODUCCIÓN'].h;
        if (textosInterfase['INTRODUCCIÓN'].i) {
        introHTML += ` <strong>${textosInterfase['INTRODUCCIÓN'].i}</strong>.`;
        }
        document.getElementById('b5-lead').innerHTML = introHTML;
    }

    if (textosInterfase['FECHA DE CORTE']) document.getElementById('b5-date').innerText = `Corte al ${textosInterfase['FECHA DE CORTE'].h.toLowerCase()}`;

    // 2. Tarjeta Histórico (Lado izquierdo)
    if (textosInterfase['TARJETA 1 PRIMERA LÍNEA']) document.getElementById('b5-t1-tag').innerText = textosInterfase['TARJETA 1 PRIMERA LÍNEA'].h;
    if (textosInterfase['TARJETA 1 SECUNDARIA']) document.getElementById('b5-t1-desc').innerText = textosInterfase['TARJETA 1 SECUNDARIA'].h;
    
    let totalEgresados = nivelesEgresados.reduce((sum, item) => sum + item.valor, 0);
    document.getElementById('b5-total-num').setAttribute('data-target', totalEgresados);

    let htmlBreakdown = '';
    const coloresUAA = { 'licenciatura': 'var(--uaa-navy)', 'media': 'var(--uaa-teal)', 'posgrado': 'var(--uaa-magenta)' };

    nivelesEgresados.forEach(niv => {
        let color = 'var(--uaa-navy)';
        const nomLow = niv.nombre.toLowerCase();
        if (nomLow.includes('media') || nomLow.includes('bachillerato')) color = coloresUAA['media'];
        if (nomLow.includes('posgrado') || nomLow.includes('maestría') || nomLow.includes('doctorado')) color = coloresUAA['posgrado'];

        htmlBreakdown += `
        <div class="breakdown-item">
            <div class="breakdown-header"><span class="breakdown-label">${niv.nombre}</span><span class="breakdown-val counter" data-target="${niv.valor}">0</span></div>
            <div class="progress-track"><div class="progress-fill" data-value="${niv.valor}" data-total="${totalEgresados}" style="background: ${color}; width: 0%;"></div></div>
        </div>`;
    });
    document.getElementById('b5-breakdown').innerHTML = htmlBreakdown;

    // 3. Tarjeta Gráfica de Dona (Lado derecho)
    if (textosInterfase['TARJETA 2 PRIMERA LÍNEA']) document.getElementById('b5-t2-tag').innerText = textosInterfase['TARJETA 2 PRIMERA LÍNEA'].h;

    const R = 70;
    const CIRCUMFERENCE = 2 * Math.PI * R;
    let htmlDonutLabels = '';
    let htmlDonutSegments = '';
    let currentOffset = 0;
    
    const posiciones = [
        'bottom: 10%; right: -25%;',
        'top: 25%; left: -25%;',
        'top: -10%; left: 30%;'
    ];

    nivelesEgresados.sort((a, b) => b.valor - a.valor);

    nivelesEgresados.forEach((niv, index) => {
        let color = 'var(--uaa-navy)';
        const nomLow = niv.nombre.toLowerCase();
        if (nomLow.includes('media') || nomLow.includes('bachillerato')) color = 'var(--uaa-teal)';
        if (nomLow.includes('posgrado') || nomLow.includes('maestría') || nomLow.includes('doctorado')) color = 'var(--uaa-magenta)';

        const pct = totalEgresados > 0 ? (niv.valor / totalEgresados) * 100 : 0;
        const dashLength = (pct / 100) * CIRCUMFERENCE;
        const coord = posiciones[index] || posiciones[0];

        htmlDonutLabels += `
        <div class="floating-label" style="${coord} transition-delay: 0.${index+1}s;">
            <span class="dot" style="background: ${color};"></span>
            <span>${niv.nombre} <strong>${pct.toFixed(1)}%</strong></span>
        </div>`;
        
        htmlDonutSegments += `
        <circle class="donut-segment" cx="90" cy="90" r="${R}" stroke="${color}" data-dash="${dashLength} ${CIRCUMFERENCE}" stroke-dashoffset="${-currentOffset}" style="stroke-dasharray: 0 440;"></circle>`;
        
        currentOffset += dashLength;
    });

    const htmlDonutFull = `
        <div class="donut-wrapper">
        ${htmlDonutLabels}
        <svg class="donut-svg" viewBox="0 0 180 180">
            <circle class="donut-circle-bg" cx="90" cy="90" r="70" />
            ${htmlDonutSegments}
        </svg>
        <div class="donut-center-info">
            <div class="donut-center-total">100%</div>
            <div class="donut-center-label">Total</div>
        </div>
        </div>
    `;
    document.getElementById('b5-donut-container').innerHTML = htmlDonutFull;

    // =========================================================================
    // DETONADOR DE ANIMACIONES
    // =========================================================================
    setTimeout(() => {
        const cardsContainer = document.getElementById('bloque-05-cards');
        if(!cardsContainer) return;

        cardsContainer.querySelectorAll('.counter').forEach(c => {
        const target = parseFloat(c.getAttribute('data-target') || 0);
        animarDato(c, target);
        });

        cardsContainer.querySelectorAll('.progress-fill').forEach(f => {
        const val = parseFloat(f.getAttribute('data-value') || 0);
        const tot = parseFloat(f.getAttribute('data-total') || 1);
        if(tot > 0) f.style.width = (val / tot * 100) + '%';
        });

        cardsContainer.querySelectorAll('.donut-segment').forEach(seg => { 
        seg.style.strokeDasharray = seg.getAttribute('data-dash'); 
        });
    }, 50);

    })
    .catch(error => {
    console.error('Detalle del error en Bloque 05:', error);
    });
});

// =========================================================================
// AUTOMATIZACIÓN BLOQUE 06: PERSONAL ACADÉMICO
// =========================================================================
window.addEventListener('DOMContentLoaded', () => {
fetchSheetData('06_Personal_Academico').then(data => {
    
    let textosInterfase = {};
    let contratacion = [];
    let sniiGen = { mujeres: 0, hombres: 0 };
    let sniiNivel = [];
    let prodepGen = { mujeres: 0, hombres: 0 };
    let cuerpos = {};

    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        if (!row || row.length === 0) continue;

        // Textos de configuración (Col G, H e I)
        const keyTexto = (row[6] || '').trim();
        if (keyTexto) {
        textosInterfase[keyTexto] = { h: (row[7] || '').trim(), i: (row[8] || '').trim() };
        }

        // Datos de la tabla (Col A, B, C)
        let cat = (row[0] || '').trim();
        let ind = (row[1] || '').trim();
        let val = cleanNum(row[2]);

        if (cat && cat !== 'Categoría') {
        if (cat === 'Tipo de Contratación') contratacion.push({ nombre: ind, valor: val });
        else if (cat === 'Investigadores (SNII) - Género') {
            if (ind.toLowerCase() === 'mujeres') sniiGen.mujeres = val;
            if (ind.toLowerCase() === 'hombres') sniiGen.hombres = val;
        }
        else if (cat === 'Investigadores (SNII) - Nivel') sniiNivel.push({ nombre: ind, valor: val });
        else if (cat === 'Perfil PRODEP') {
            if (ind.toLowerCase() === 'mujeres') prodepGen.mujeres = val;
            if (ind.toLowerCase() === 'hombres') prodepGen.hombres = val;
        }
        else if (cat === 'Cuerpos Académicos') cuerpos[ind] = val;
        }
    }

    const b6 = document.getElementById('bloque-06-cards');
    if (!b6) return;

    // 1. Textos Generales
    if (textosInterfase['PRIMERA LÍNEA']) document.getElementById('b6-tag-text').innerText = textosInterfase['PRIMERA LÍNEA'].h;
    if (textosInterfase['LÍNEA PRINCIPAL']) document.getElementById('b6-title').innerText = textosInterfase['LÍNEA PRINCIPAL'].h;
    
    if (textosInterfase['INTRODUCCIÓN']) {
        let intro = textosInterfase['INTRODUCCIÓN'].h;
        if (textosInterfase['INTRODUCCIÓN'].i) intro += ` <strong>${textosInterfase['INTRODUCCIÓN'].i}</strong>.`;
        document.getElementById('b6-lead').innerHTML = intro;
    }
    if (textosInterfase['FECHA DE CORTE']) document.getElementById('b6-date').innerText = `Corte al ${textosInterfase['FECHA DE CORTE'].h.toLowerCase()}`;

    // 2. Tarjeta 1: Plantilla Total
    if (textosInterfase['TARJETA 1 PRIMERA LÍNEA']) document.getElementById('b6-t1-tag').innerText = textosInterfase['TARJETA 1 PRIMERA LÍNEA'].h;
    if (textosInterfase['TARJETA 1 SECUNDARIA']) document.getElementById('b6-t1-desc').innerText = textosInterfase['TARJETA 1 SECUNDARIA'].h;

    let totalPlantilla = contratacion.reduce((sum, item) => sum + item.valor, 0);
    let tecnicosAc = contratacion.find(c => c.nombre.includes('Técnicos'))?.valor || 0;
    let profesores = totalPlantilla - tecnicosAc;

    document.getElementById('b6-total-plantilla').setAttribute('data-target', totalPlantilla);
    document.getElementById('b6-plantilla-texto').innerHTML = `Conformado por <strong>${profesores.toLocaleString('es-MX')} profesores</strong> de distintas dedicaciones y <strong>${tecnicosAc.toLocaleString('es-MX')} Técnicos Académicos</strong> especializados.`;

    // 3. Tarjeta 2: Tipos de Contratación (Barras)
    if (textosInterfase['TARJETA 2 PRIMERA LÍNEA']) document.getElementById('b6-t2-tag').innerText = textosInterfase['TARJETA 2 PRIMERA LÍNEA'].h;
    
    let htmlContratacion = '';
    const coloresContratacion = ['var(--uaa-navy)', 'var(--uaa-teal)', 'var(--uaa-gold-vibrant)', 'var(--uaa-magenta)'];
    
    contratacion.forEach((c, index) => {
        const pct = totalPlantilla > 0 ? (c.valor / totalPlantilla) * 100 : 0;
        const color = coloresContratacion[index % coloresContratacion.length];
        htmlContratacion += `
        <div class="bar-item">
            <div class="bar-labels"><span class="bar-name">${c.nombre} (${c.valor.toLocaleString('es-MX')})</span><span class="bar-value">${pct.toFixed(1)}%</span></div>
            <div class="progress-track"><div class="progress-fill" style="width: 0%; background: ${color};" data-percent="${pct.toFixed(1)}%"></div></div>
        </div>`;
    });
    document.getElementById('b6-bars-contratacion').innerHTML = htmlContratacion;

    // 4. Tarjeta 3: SNII
    if (textosInterfase['TARJETA 3 PRIMERA LÍNEA']) document.getElementById('b6-t3-tag').innerText = textosInterfase['TARJETA 3 PRIMERA LÍNEA'].h;
    if (textosInterfase['TARJETA 3 SECUNDARIA']) document.getElementById('b6-t3-sec').innerText = textosInterfase['TARJETA 3 SECUNDARIA'].h;
    
    let totalSnii = sniiNivel.reduce((sum, item) => sum + item.valor, 0);
    document.getElementById('b6-snii-total').setAttribute('data-target', totalSnii);
    document.getElementById('b6-snii-texto').innerHTML = `Profesores con distinción SNII (<strong>${sniiGen.mujeres} mujeres</strong> y <strong>${sniiGen.hombres} hombres</strong>).`;

    let htmlSnii = '';
    sniiNivel.forEach(n => {
        htmlSnii += `
        <div class="breakdown-item">
            <div class="breakdown-header"><span class="breakdown-label">${n.nombre}</span><span class="breakdown-val counter" data-target="${n.valor}">0</span></div>
            <div class="progress-track"><div class="progress-fill" data-value="${n.valor}" data-total="${totalSnii}" style="background: #565D6D; width: 0%;"></div></div>
        </div>`;
    });
    document.getElementById('b6-snii-breakdown').innerHTML = htmlSnii;

    // 5. Tarjeta 4: PRODEP
    if (textosInterfase['TARJETA 4 PRIMERA LÍNEA']) document.getElementById('b6-t4-tag').innerText = textosInterfase['TARJETA 4 PRIMERA LÍNEA'].h;
    if (textosInterfase['TARJETA 4 SECUNDARIA']) document.getElementById('b6-t4-sec').innerText = textosInterfase['TARJETA 4 SECUNDARIA'].h;

    let totalProdep = prodepGen.mujeres + prodepGen.hombres;
    document.getElementById('b6-prodep-total').setAttribute('data-target', totalProdep);
    document.getElementById('b6-prodep-texto').innerHTML = `Fortalecimiento de las funciones sustantivas de la institución, integrado por <strong>${prodepGen.mujeres} mujeres</strong> y <strong>${prodepGen.hombres} hombres</strong> con reconocimiento vigente.`;

    // 6. Tarjeta 5: Cuerpos Académicos
    if (textosInterfase['TARJETA 5 PRIMERA LÍNEA']) document.getElementById('b6-t5-tag').innerText = textosInterfase['TARJETA 5 PRIMERA LÍNEA'].h;
    if (textosInterfase['TARJETA 5 SECUNDARIA']) document.getElementById('b6-t5-sec').innerText = textosInterfase['TARJETA 5 SECUNDARIA'].h;

    let consolidados = cuerpos['Consolidados'] || 0;
    let consolidacion = cuerpos['En Consolidación'] || 0;
    let formacion = cuerpos['En Formación'] || 0;
    let vigentes = consolidados + consolidacion + formacion;

    document.getElementById('b6-ca-miembros').setAttribute('data-target', cuerpos['Miembros Adscritos'] || 0);
    document.getElementById('b6-ca-consolidados').setAttribute('data-target', consolidados);
    document.getElementById('b6-ca-consolidacion').setAttribute('data-target', consolidacion);
    document.getElementById('b6-ca-formacion').setAttribute('data-target', formacion);
    document.getElementById('b6-ca-vigentes').setAttribute('data-target', vigentes);

    // =========================================================================
    // DETONADOR DE ANIMACIONES
    // =========================================================================
    setTimeout(() => {
        const cardsContainer = document.getElementById('bloque-06-cards');
        if(!cardsContainer) return;

        cardsContainer.querySelectorAll('.counter').forEach(c => {
        const target = parseFloat(c.getAttribute('data-target') || 0);
        animarDato(c, target);
        });

        cardsContainer.querySelectorAll('.progress-fill').forEach(f => {
        if (f.hasAttribute('data-percent')) {
            f.style.width = f.getAttribute('data-percent');
        } else {
            const val = parseFloat(f.getAttribute('data-value') || 0);
            const tot = parseFloat(f.getAttribute('data-total') || 1);
            if(tot > 0) f.style.width = (val / tot * 100) + '%';
        }
        });
    }, 50);

    })
    .catch(error => {
    console.error('Detalle del error en Bloque 06:', error);
    });
});

// =========================================================================
// AUTOMATIZACIÓN BLOQUE 07: PROYECTOS DE INVESTIGACIÓN VIGENTES
// =========================================================================
window.addEventListener('DOMContentLoaded', () => {
fetchSheetData('07_Proyectos_Inv').then(data => {

    let textosInterfase = {};
    let totalProyectos = 0;
    let redColaboracion = 0;
    let formacionEstudiantil = 0;
    let impactoSocial = 0;

    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        if (!row || row.length === 0) continue;

        // Mapeo Textos Naranjas (Col G y H)
        const keyTexto = (row[6] || '').trim();
        if (keyTexto) {
        textosInterfase[keyTexto] = { h: (row[7] || '').trim() };
        }

        // Mapeo Datos de Tabla (Col A y B)
        let indicador = (row[0] || '').trim();
        let valor = cleanNum(row[1]);

        if (indicador.includes('Total de Proyectos')) totalProyectos = valor;
        else if (indicador.includes('Red de Colaboración')) redColaboracion = valor;
        else if (indicador.includes('Formación Estudiantil')) formacionEstudiantil = valor;
        else if (indicador.includes('Impacto Social')) impactoSocial = valor;
    }

    const b7 = document.getElementById('bloque-07-cards');
    if (!b7) return;

    // 1. Textos Generales
    if (textosInterfase['PRIMERA LÍNEA']) document.getElementById('b7-tag-text').innerText = textosInterfase['PRIMERA LÍNEA'].h;
    if (textosInterfase['LÍNEA PRINCIPAL']) document.getElementById('b7-title').innerText = textosInterfase['LÍNEA PRINCIPAL'].h;
    if (textosInterfase['INTRODUCCIÓN']) document.getElementById('b7-lead').innerText = textosInterfase['INTRODUCCIÓN'].h;
    if (textosInterfase['FECHA DE CORTE']) document.getElementById('b7-date').innerText = `Corte al ${textosInterfase['FECHA DE CORTE'].h.toLowerCase()}`;

    // 2. Tarjeta Panorama General
    if (textosInterfase['TARJETA 1 PRIMERA LÍNEA']) document.getElementById('b7-t1-tag').innerText = textosInterfase['TARJETA 1 PRIMERA LÍNEA'].h;
    if (textosInterfase['TARJETA 1 SECUNDARIA']) document.getElementById('b7-t1-desc').innerText = textosInterfase['TARJETA 1 SECUNDARIA'].h;

    document.getElementById('b7-total-proyectos').setAttribute('data-target', totalProyectos);

    // 3. Tarjetas secundarias y cálculo autónomo de porcentajes
    document.getElementById('b7-red-num').setAttribute('data-target', redColaboracion);
    const pctRed = totalProyectos > 0 ? ((redColaboracion / totalProyectos) * 100).toFixed(1) : 0;
    document.getElementById('b7-red-pct').innerText = `${pctRed}%`;

    document.getElementById('b7-formacion-num').setAttribute('data-target', formacionEstudiantil);
    const pctFormacion = totalProyectos > 0 ? ((formacionEstudiantil / totalProyectos) * 100).toFixed(1) : 0;
    document.getElementById('b7-formacion-pct').innerText = `${pctFormacion}%`;

    document.getElementById('b7-impacto-num').setAttribute('data-target', impactoSocial);
    const pctImpacto = totalProyectos > 0 ? ((impactoSocial / totalProyectos) * 100).toFixed(1) : 0;
    document.getElementById('b7-impacto-pct').innerText = `${pctImpacto}%`;

    // =========================================================================
    // DETONADOR DE ANIMACIONES
    // =========================================================================
    setTimeout(() => {
        const cardsContainer = document.getElementById('bloque-07-cards');
        if(!cardsContainer) return;

        cardsContainer.querySelectorAll('.counter').forEach(c => {
        const target = parseFloat(c.getAttribute('data-target') || 0);
        animarDato(c, target);
        });
    }, 50);

    })
    .catch(error => {
    console.error('Detalle del error en Bloque 07:', error);
    });
});

// =========================================================================
// AUTOMATIZACIÓN BLOQUE 08: PUBLICACIONES CIENTÍFICAS
// =========================================================================
window.addEventListener('DOMContentLoaded', () => {
fetchSheetData('08_Publicaciones').then(data => {

    let textosInterfase = {};
    let publicaciones = [];
    let totalNacional = 0;
    let totalInternacional = 0;

    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        if (!row || row.length === 0) continue;

        // Mapeo de Textos (Col G y H)
        const keyTexto = (row[6] || '').trim();
        if (keyTexto) {
        textosInterfase[keyTexto] = { h: (row[7] || '').trim() };
        }

        // Mapeo de Datos de la Tabla (Col A, B, C)
        let tipo = (row[0] || '').trim();
        let nac = cleanNum(row[1]);
        let inter = cleanNum(row[2]);

        if (tipo && tipo !== 'Tipo de Publicación') {
        let filaTotal = nac + inter;
        publicaciones.push({ tipo, nac, inter, filaTotal });
        
        // Acumuladores para el total final
        totalNacional += nac;
        totalInternacional += inter;
        }
    }

    const b8 = document.getElementById('bloque-08-cards');
    if (!b8) return;

    let granTotal = totalNacional + totalInternacional;

    // 1. Textos Generales
    if (textosInterfase['PRIMERA LÍNEA']) document.getElementById('b8-tag-text').innerText = textosInterfase['PRIMERA LÍNEA'].h;
    if (textosInterfase['LÍNEA PRINCIPAL']) document.getElementById('b8-title').innerText = textosInterfase['LÍNEA PRINCIPAL'].h;
    if (textosInterfase['INTRODUCCIÓN']) document.getElementById('b8-lead').innerText = textosInterfase['INTRODUCCIÓN'].h;
    if (textosInterfase['FECHA DE CORTE']) document.getElementById('b8-date').innerText = `Corte al ${textosInterfase['FECHA DE CORTE'].h.toLowerCase()}`;
    if (textosInterfase['TARJETA 1 PRIMERA LÍNEA']) document.getElementById('b8-t1-tag').innerText = textosInterfase['TARJETA 1 PRIMERA LÍNEA'].h;
    if (textosInterfase['TARJETA 1 SECUNDARIA']) document.getElementById('b8-t1-sec').innerText = textosInterfase['TARJETA 1 SECUNDARIA'].h;

    // 2. Métrica Principal
    document.getElementById('b8-total-pub').setAttribute('data-target', granTotal);

    // 3. Construir la Tabla Dinámica
    let htmlTabla = '';
    publicaciones.forEach((pub, index) => {
        // Si es la última fila antes del total, le ponemos la línea inferior gruesa de tu diseño original
        let estiloFila = (index === publicaciones.length - 1) ? 'style="border-bottom: 2px solid rgba(22,48,114,0.15);"' : '';

        htmlTabla += `
        <tr ${estiloFila}>
            <td>${pub.tipo}</td>
            <td class="counter" data-target="${pub.nac}">0</td>
            <td class="counter" data-target="${pub.inter}">0</td>
            <td class="counter" data-target="${pub.filaTotal}">0</td>
        </tr>
        `;
    });

    // 4. Agregar la fila del TOTAL
    htmlTabla += `
        <tr class="data-table-total">
        <td style="font-family: var(--font-display); text-transform: uppercase; font-size: 12px; letter-spacing: 0.05em;">TOTAL</td>
        <td class="counter" data-target="${totalNacional}">0</td>
        <td class="counter" data-target="${totalInternacional}">0</td>
        <td class="counter" data-target="${granTotal}">0</td>
        </tr>
    `;

    document.getElementById('b8-table-body').innerHTML = htmlTabla;

    // =========================================================================
    // DETONADOR DE ANIMACIONES
    // =========================================================================
    setTimeout(() => {
        const cardsContainer = document.getElementById('bloque-08-cards');
        if(!cardsContainer) return;

        // Detonar contadores (incluye los de la tabla que acaban de inyectarse)
        cardsContainer.querySelectorAll('.counter').forEach(c => {
        const target = parseFloat(c.getAttribute('data-target') || 0);
        animarDato(c, target);
        });
    }, 50);

    })
    .catch(error => {
    console.error('Detalle del error en Bloque 08:', error);
    });
});

// =========================================================================
// AUTOMATIZACIÓN BLOQUE 09: PATENTES INSTITUCIONALES
// =========================================================================
window.addEventListener('DOMContentLoaded', () => {
fetchSheetData('09_Patentes').then(data => {

    let textosInterfase = {};
    let otorgadas = 0;
    let enProceso = 0;

    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        if (!row || row.length === 0) continue;

        // Mapeo Textos Naranjas (Col G, H, I)
        const keyTexto = (row[6] || '').trim();
        if (keyTexto) {
        textosInterfase[keyTexto] = { 
            h: (row[7] || '').trim(),
            i: (row[8] || '').trim() 
        };
        }

        // Mapeo Datos Tabla (Col A y B)
        let estatus = (row[0] || '').trim();
        let cantidad = cleanNum(row[1]);

        if (estatus.includes('Otorgadas')) otorgadas = cantidad;
        else if (estatus.includes('En Proceso')) enProceso = cantidad;
    }

    const b9 = document.getElementById('bloque-09-cards');
    if (!b9) return;

    let totalPatentes = otorgadas + enProceso;

    // 1. Textos Generales
    if (textosInterfase['PRIMERA LÍNEA']) document.getElementById('b9-tag-text').innerText = textosInterfase['PRIMERA LÍNEA'].h;
    if (textosInterfase['LÍNEA PRINCIPAL']) document.getElementById('b9-title').innerText = textosInterfase['LÍNEA PRINCIPAL'].h;
    if (textosInterfase['INTRODUCCIÓN']) document.getElementById('b9-lead').innerText = textosInterfase['INTRODUCCIÓN'].h;
    if (textosInterfase['FECHA DE CORTE']) document.getElementById('b9-date').innerText = `Corte al ${textosInterfase['FECHA DE CORTE'].h.toLowerCase()}`;
    
    if (textosInterfase['LINK SUPERIOR']) {
        const linkSup = document.getElementById('b9-link-sup');
        linkSup.innerHTML = `${textosInterfase['LINK SUPERIOR'].h} &#8599;`;
        linkSup.href = textosInterfase['LINK SUPERIOR'].i;
    }

    // 2. Tarjeta Patentes
    if (textosInterfase['TARJETA 1 PRIMERA LÍNEA']) document.getElementById('b9-t1-tag').innerText = textosInterfase['TARJETA 1 PRIMERA LÍNEA'].h;
    if (textosInterfase['TARJETA 1 SECUNDARIA']) document.getElementById('b9-t1-sec').innerText = textosInterfase['TARJETA 1 SECUNDARIA'].h;
    
    if (textosInterfase['TARJETA 1 LINK']) {
        const t1Link = document.getElementById('b9-t1-link');
        t1Link.innerHTML = `${textosInterfase['TARJETA 1 LINK'].h} &#8599;`;
        t1Link.href = textosInterfase['TARJETA 1 LINK'].i;
    }

    document.getElementById('b9-total-num').setAttribute('data-target', totalPatentes);
    document.getElementById('b9-otorgadas-num').setAttribute('data-target', otorgadas);
    document.getElementById('b9-proceso-num').setAttribute('data-target', enProceso);

    // =========================================================================
    // DETONADOR DE ANIMACIONES
    // =========================================================================
    setTimeout(() => {
        const cardsContainer = document.getElementById('bloque-09-cards');
        if(!cardsContainer) return;

        cardsContainer.querySelectorAll('.counter').forEach(c => {
        const target = parseFloat(c.getAttribute('data-target') || 0);
        animarDato(c, target);
        });
    }, 50);

    })
    .catch(error => {
    console.error('Detalle del error en Bloque 09:', error);
    });
});

// =========================================================================
// AUTOMATIZACIÓN BLOQUE 10: PRODUCCIÓN EDITORIAL
// =========================================================================
window.addEventListener('DOMContentLoaded', () => {
fetchSheetData('10_Editorial').then(data => {

    let textosInterfase = {};
    let impresas = 0;
    let digitales = 0;
    let revistas = 0;

    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        if (!row || row.length === 0) continue;

        // Mapeo Textos Naranjas (Col G, H)
        const keyTexto = (row[6] || '').trim();
        if (keyTexto) {
        textosInterfase[keyTexto] = { h: (row[7] || '').trim() };
        }

        // Mapeo Datos Tabla (Col A y B)
        let tipoObra = (row[0] || '').trim();
        let cantidad = cleanNum(row[1]);

        if (tipoObra.includes('Impresas')) impresas = cantidad;
        else if (tipoObra.includes('Digitales')) digitales = cantidad;
        else if (tipoObra.includes('Revistas')) revistas = cantidad;
    }

    const b10 = document.getElementById('bloque-10-cards');
    if (!b10) return;

    // 1. Textos Generales
    if (textosInterfase['PRIMERA LÍNEA']) document.getElementById('b10-tag-text').innerText = textosInterfase['PRIMERA LÍNEA'].h;
    if (textosInterfase['LÍNEA PRINCIPAL']) document.getElementById('b10-title').innerText = textosInterfase['LÍNEA PRINCIPAL'].h;
    if (textosInterfase['INTRODUCCIÓN']) document.getElementById('b10-lead').innerText = textosInterfase['INTRODUCCIÓN'].h;
    
    let fechaCompleta = textosInterfase['FECHA DE CORTE'] ? textosInterfase['FECHA DE CORTE'].h : '';
    document.getElementById('b10-date').innerText = `Corte al ${fechaCompleta.toLowerCase()}`;
    
    // Extraemos solo el número (año) para inyectarlo en los textos inferiores
    let anioCorte = fechaCompleta.replace(/[^\d]/g, '') || '2025';

    // 2. Tarjetas Individuales
    document.getElementById('b10-impresas-num').setAttribute('data-target', impresas);
    document.getElementById('b10-impresas-desc').innerHTML = `Obras editoriales impresas publicadas durante el ejercicio <strong>${anioCorte}</strong>.`;

    document.getElementById('b10-digitales-num').setAttribute('data-target', digitales);
    document.getElementById('b10-digitales-desc').innerHTML = `Obras editoriales digitales publicadas durante el ejercicio <strong>${anioCorte}</strong>.`;

    document.getElementById('b10-revistas-num').setAttribute('data-target', revistas);
    document.getElementById('b10-revistas-desc').innerHTML = `Revistas publicadas en diferentes formatos y periodicidad durante el ejercicio <strong>${anioCorte}</strong>.`;

    // =========================================================================
    // DETONADOR DE ANIMACIONES
    // =========================================================================
    setTimeout(() => {
        const cardsContainer = document.getElementById('bloque-10-cards');
        if(!cardsContainer) return;

        cardsContainer.querySelectorAll('.counter').forEach(c => {
        const target = parseFloat(c.getAttribute('data-target') || 0);
        animarDato(c, target);
        });
    }, 50);

    })
    .catch(error => {
    console.error('Detalle del error en Bloque 10:', error);
    });
});

// =========================================================================
// AUTOMATIZACIÓN BLOQUE 11: INFRAESTRUCTURA UNIVERSITARIA
// =========================================================================
window.addEventListener('DOMContentLoaded', () => {
fetchSheetData('11_Infraestructura').then(data => {

    let textosInterfase = {};
    let totalMetros = 0;

    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        if (!row || row.length === 0) continue;

        // Mapeo Textos Naranjas (Col G, H, I)
        const keyTexto = (row[6] || '').trim();
        if (keyTexto) {
        textosInterfase[keyTexto] = { 
            h: (row[7] || '').trim(),
            i: (row[8] || '').trim() 
        };
        }

        // Mapeo Datos Tabla (Col A y B)
        let tipoEspacio = (row[0] || '').trim();
        let m2 = cleanNum(row[1]);

        if (tipoEspacio && tipoEspacio !== 'Tipo de Espacio') {
        totalMetros += m2;
        
        // Asignación inteligente por palabras clave a las 8 tarjetas
        const tipoLower = tipoEspacio.toLowerCase();
        
        if (tipoLower.includes('académicos')) {
            document.getElementById('b11-lbl-academicos').innerText = tipoEspacio;
            document.getElementById('b11-num-academicos').setAttribute('data-target', m2);
        } 
        else if (tipoLower.includes('docencia')) {
            document.getElementById('b11-lbl-docencia').innerText = tipoEspacio;
            document.getElementById('b11-num-docencia').setAttribute('data-target', m2);
        }
        else if (tipoLower.includes('otros')) {
            document.getElementById('b11-lbl-otros').innerText = tipoEspacio;
            document.getElementById('b11-num-otros').setAttribute('data-target', m2);
        }
        else if (tipoLower.includes('cómputo')) {
            document.getElementById('b11-lbl-computo').innerText = tipoEspacio;
            document.getElementById('b11-num-computo').setAttribute('data-target', m2);
        }
        else if (tipoLower.includes('laboratorios de especialidad')) {
            document.getElementById('b11-lbl-labespecialidad').innerText = tipoEspacio;
            document.getElementById('b11-num-labespecialidad').setAttribute('data-target', m2);
        }
        else if (tipoLower.includes('talleres')) {
            document.getElementById('b11-lbl-talleres').innerText = tipoEspacio;
            document.getElementById('b11-num-talleres').setAttribute('data-target', m2);
        }
        else if (tipoLower.includes('administrativos')) {
            document.getElementById('b11-lbl-admin').innerText = tipoEspacio;
            document.getElementById('b11-num-admin').setAttribute('data-target', m2);
        }
        else if (tipoLower.includes('estacionamiento')) {
            document.getElementById('b11-lbl-estacionamientos').innerText = tipoEspacio;
            document.getElementById('b11-num-estacionamientos').setAttribute('data-target', m2);
        }
        }
    }

    const b11 = document.getElementById('bloque-11-cards');
    if (!b11) return;

    // 1. Textos Generales
    if (textosInterfase['PRIMERA LÍNEA']) document.getElementById('b11-tag-text').innerText = textosInterfase['PRIMERA LÍNEA'].h;
    if (textosInterfase['LÍNEA PRINCIPAL']) document.getElementById('b11-title').innerText = textosInterfase['LÍNEA PRINCIPAL'].h;
    
    // Armamos el texto combinando la Columna H, la suma matemática y la Columna I
    if (textosInterfase['INTRODUCCIÓN']) {
        let introH = textosInterfase['INTRODUCCIÓN'].h;
        let introI = textosInterfase['INTRODUCCIÓN'].i ? textosInterfase['INTRODUCCIÓN'].i : '';
        document.getElementById('b11-lead').innerHTML = `${introH} <strong>${totalMetros.toLocaleString('es-MX')} metros cuadrados</strong> ${introI}`;
    }
    
    if (textosInterfase['FECHA DE CORTE']) document.getElementById('b11-date').innerText = `Corte al ${textosInterfase['FECHA DE CORTE'].h.toLowerCase()}`;
    
    if (textosInterfase['LINK SUPERIOR']) {
        const linkSup = document.getElementById('b11-link-sup');
        linkSup.innerHTML = `${textosInterfase['LINK SUPERIOR'].h} &#8599;`;
        linkSup.href = textosInterfase['LINK SUPERIOR'].i;
    }

    // =========================================================================
    // DETONADOR DE ANIMACIONES
    // =========================================================================
    setTimeout(() => {
        const cardsContainer = document.getElementById('bloque-11-cards');
        if(!cardsContainer) return;

        cardsContainer.querySelectorAll('.counter').forEach(c => {
        const target = parseFloat(c.getAttribute('data-target') || 0);
        animarDato(c, target);
        });
    }, 50);

    })
    .catch(error => {
    console.error('Detalle del error en Bloque 11:', error);
    });
});

// =========================================================================
// AUTOMATIZACIÓN BLOQUE 12: DIRECTORIO DE ESPACIOS
// =========================================================================
window.addEventListener('DOMContentLoaded', () => {
fetchSheetData('12_Directorio').then(data => {

    let textosInterfase = {};
    
    // Diccionario para almacenar el HTML acumulado por categoría
    let listasHTML = {
        bibliotecas: '',
        auditorios: '',
        deportivos: '',
        esparcimiento: '',
        culturales: '',
        talleres: '',
        laboratorios: '',
        otros: ''
    };

    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        if (!row || row.length === 0) continue;

        // Mapeo Textos Naranjas (Por si en tu hoja los tienes en Col G y H)
        const keyTexto = (row[6] || '').trim();
        if (keyTexto) {
        textosInterfase[keyTexto] = { h: (row[7] || '').trim() };
        }

        // Mapeo Directorio (Col A, B, C)
        let categoria = (row[0] || '').trim();
        let nombre = (row[1] || '').trim();
        let enlace = (row[2] || '').trim();

        if (categoria && categoria !== 'Categoría' && nombre) {
        let catLower = categoria.toLowerCase();
        
        // Construimos la línea de HTML (el <li> con su icono y link)
        let itemHTML = `<li class="directory-item"><svg class="directory-icon"><use href="#dir-icon"></use></svg><a href="${enlace}" target="_blank" class="directory-link">${nombre}</a></li>`;

        // Lo agrupamos en la lista que le corresponde
        if (catLower.includes('bibliotecas')) listasHTML.bibliotecas += itemHTML;
        else if (catLower.includes('auditorios')) listasHTML.auditorios += itemHTML;
        else if (catLower.includes('deportivos')) listasHTML.deportivos += itemHTML;
        else if (catLower.includes('esparcimiento')) listasHTML.esparcimiento += itemHTML;
        else if (catLower.includes('culturales')) listasHTML.culturales += itemHTML;
        else if (catLower.includes('talleres')) listasHTML.talleres += itemHTML;
        else if (catLower.includes('laboratorios')) listasHTML.laboratorios += itemHTML;
        else if (catLower.includes('otros')) listasHTML.otros += itemHTML;
        }
    }

    const b12 = document.getElementById('bloque-12-cards');
    if (!b12) return;

    // 1. Textos Generales (si existen en el Sheets, si no, respeta el HTML)
    if (textosInterfase['PRIMERA LÍNEA']) document.getElementById('b12-tag-text').innerText = textosInterfase['PRIMERA LÍNEA'].h;
    if (textosInterfase['LÍNEA PRINCIPAL']) document.getElementById('b12-title').innerText = textosInterfase['LÍNEA PRINCIPAL'].h;
    if (textosInterfase['INTRODUCCIÓN']) document.getElementById('b12-lead').innerHTML = textosInterfase['INTRODUCCIÓN'].h;

    // 2. Inyección de las listas armadas
    if (document.getElementById('b12-list-bibliotecas')) document.getElementById('b12-list-bibliotecas').innerHTML = listasHTML.bibliotecas;
    if (document.getElementById('b12-list-auditorios')) document.getElementById('b12-list-auditorios').innerHTML = listasHTML.auditorios;
    if (document.getElementById('b12-list-deportivos')) document.getElementById('b12-list-deportivos').innerHTML = listasHTML.deportivos;
    if (document.getElementById('b12-list-esparcimiento')) document.getElementById('b12-list-esparcimiento').innerHTML = listasHTML.esparcimiento;
    if (document.getElementById('b12-list-culturales')) document.getElementById('b12-list-culturales').innerHTML = listasHTML.culturales;
    if (document.getElementById('b12-list-talleres')) document.getElementById('b12-list-talleres').innerHTML = listasHTML.talleres;
    if (document.getElementById('b12-list-laboratorios')) document.getElementById('b12-list-laboratorios').innerHTML = listasHTML.laboratorios;
    if (document.getElementById('b12-list-otros')) document.getElementById('b12-list-otros').innerHTML = listasHTML.otros;

    })
    .catch(error => {
    console.error('Detalle del error en Bloque 12:', error);
    });
}); 

// =========================================================================
// AUTOMATIZACIÓN BLOQUE 13: PERSONAL ADMINISTRATIVO
// =========================================================================
window.addEventListener('DOMContentLoaded', () => {
fetchSheetData('13_Personal_Administrativo').then(data => {

    let textosInterfase = {};
    
    let sindH = 0, sindM = 0;
    let confH = 0, confM = 0;

    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        if (!row || row.length === 0) continue;

        // Mapeo Textos (Col G, H, I)
        const keyTexto = (row[6] || '').trim();
        if (keyTexto) {
        textosInterfase[keyTexto] = { h: (row[7] || '').trim(), i: (row[8] || '').trim() };
        }

        // Mapeo Datos Tabla (Col A, B, C)
        let regimen = (row[0] || '').trim();
        if (regimen.includes('Sindicalizados')) {
        sindH = cleanNum(row[1]);
        sindM = cleanNum(row[2]);
        } else if (regimen.includes('De Confianza')) {
        confH = cleanNum(row[1]);
        confM = cleanNum(row[2]);
        }
    }

    const b13 = document.getElementById('bloque-13-cards');
    if (!b13) return;

    // Cálculos Matemáticos
    let sindTotal = sindH + sindM;
    let confTotal = confH + confM;
    let granTotal = sindTotal + confTotal;
    let totalHombres = sindH + confH;
    let totalMujeres = sindM + confM;

    // 1. Textos Generales
    if (textosInterfase['PRIMERA LÍNEA']) document.getElementById('b13-tag-text').innerText = textosInterfase['PRIMERA LÍNEA'].h;
    if (textosInterfase['LÍNEA PRINCIPAL']) document.getElementById('b13-title').innerText = textosInterfase['LÍNEA PRINCIPAL'].h;
    if (textosInterfase['INTRODUCCIÓN']) document.getElementById('b13-lead').innerText = textosInterfase['INTRODUCCIÓN'].h;
    if (textosInterfase['FECHA DE CORTE']) document.getElementById('b13-date').innerText = `Corte al ${textosInterfase['FECHA DE CORTE'].h.toLowerCase()}`;
    if (textosInterfase['LINK SUPERIOR']) {
        const linkSup = document.getElementById('b13-link-sup');
        linkSup.innerHTML = `${textosInterfase['LINK SUPERIOR'].h} &#8599;`;
        linkSup.href = textosInterfase['LINK SUPERIOR'].i;
    }

    // 2. Tarjeta 1: Total
    if (textosInterfase['TARJETA 1 PRIMERA LÍNEA']) document.getElementById('b13-t1-tag').innerText = textosInterfase['TARJETA 1 PRIMERA LÍNEA'].h;
    if (textosInterfase['TARJETA 1 SECUNDARIA']) document.getElementById('b13-t1-desc').innerText = textosInterfase['TARJETA 1 SECUNDARIA'].h;

    const pctMujeres = granTotal > 0 ? ((totalMujeres / granTotal) * 100).toFixed(1) : 0;
    const pctHombres = granTotal > 0 ? ((totalHombres / granTotal) * 100).toFixed(1) : 0;
    
    document.getElementById('b13-total-mujeres-txt').innerHTML = `<strong>${totalMujeres.toLocaleString('es-MX')}</strong> Mujeres (${pctMujeres}%)`;
    document.getElementById('b13-total-hombres-txt').innerHTML = `<strong>${totalHombres.toLocaleString('es-MX')}</strong> Hombres (${pctHombres}%)`;
    document.getElementById('b13-total-plantilla').setAttribute('data-target', granTotal);

    // 3. Tarjeta 2: Sindicalizados
    if (textosInterfase['TARJETA 2 PRIMERA LÍNEA']) document.getElementById('b13-t2-tag').innerText = textosInterfase['TARJETA 2 PRIMERA LÍNEA'].h;
    if (textosInterfase['TARJETA 2 SECUNDARIA']) document.getElementById('b13-t2-sec').innerText = textosInterfase['TARJETA 2 SECUNDARIA'].h;
    if (textosInterfase['TARJETA 2 CONCLUSIÓN']) document.getElementById('b13-t2-concl').innerText = textosInterfase['TARJETA 2 CONCLUSIÓN'].h;

    const sindPctTotal = granTotal > 0 ? ((sindTotal / granTotal) * 100).toFixed(1) : 0;
    document.getElementById('b13-sind-total').setAttribute('data-target', sindTotal);
    document.getElementById('b13-sind-pct').innerText = `${sindPctTotal}% del total`;

    const pctSindH = sindTotal > 0 ? ((sindH / sindTotal) * 100).toFixed(1) : 0;
    const pctSindM = sindTotal > 0 ? ((sindM / sindTotal) * 100).toFixed(1) : 0;
    document.getElementById('b13-sind-h-num').setAttribute('data-target', sindH);
    document.getElementById('b13-sind-h-pct').innerText = `${pctSindH}%`;
    document.getElementById('b13-sind-m-num').setAttribute('data-target', sindM);
    document.getElementById('b13-sind-m-pct').innerText = `${pctSindM}%`;

    // 4. Tarjeta 3: De Confianza
    if (textosInterfase['TARJETA 3 PRIMERA LÍNEA']) document.getElementById('b13-t3-tag').innerText = textosInterfase['TARJETA 3 PRIMERA LÍNEA'].h;
    if (textosInterfase['TARJETA 3 SECUNDARIA']) document.getElementById('b13-t3-sec').innerText = textosInterfase['TARJETA 3 SECUNDARIA'].h;
    if (textosInterfase['TARJETA 3 CONCLUSIÓN']) document.getElementById('b13-t3-concl').innerText = textosInterfase['TARJETA 3 CONCLUSIÓN'].h;

    const confPctTotal = granTotal > 0 ? ((confTotal / granTotal) * 100).toFixed(1) : 0;
    document.getElementById('b13-conf-total').setAttribute('data-target', confTotal);
    document.getElementById('b13-conf-pct').innerText = `${confPctTotal}% del total`;

    const pctConfH = confTotal > 0 ? ((confH / confTotal) * 100).toFixed(1) : 0;
    const pctConfM = confTotal > 0 ? ((confM / confTotal) * 100).toFixed(1) : 0;
    document.getElementById('b13-conf-h-num').setAttribute('data-target', confH);
    document.getElementById('b13-conf-h-pct').innerText = `${pctConfH}%`;
    document.getElementById('b13-conf-m-num').setAttribute('data-target', confM);
    document.getElementById('b13-conf-m-pct').innerText = `${pctConfM}%`;

    // =========================================================================
    // DETONADOR DE ANIMACIONES (INCLUYENDO LÍQUIDOS)
    // =========================================================================
    setTimeout(() => {
        const cardsContainer = document.getElementById('bloque-13-cards');
        if(!cardsContainer) return;

        // Animar todos los contadores
        cardsContainer.querySelectorAll('.counter').forEach(c => {
        const target = parseFloat(c.getAttribute('data-target') || 0);
        animarDato(c, target);
        });

        // Animar las 4 siluetas líquidas manipulando atributos SVG
        const fillSindH = (pctSindH / 100) * 200;
        document.getElementById('b13-sind-liquid-h').setAttribute('height', fillSindH);
        document.getElementById('b13-sind-liquid-h').setAttribute('y', 200 - fillSindH);

        const fillSindM = (pctSindM / 100) * 200;
        document.getElementById('b13-sind-liquid-m').setAttribute('height', fillSindM);
        document.getElementById('b13-sind-liquid-m').setAttribute('y', 200 - fillSindM);

        const fillConfH = (pctConfH / 100) * 200;
        document.getElementById('b13-conf-liquid-h').setAttribute('height', fillConfH);
        document.getElementById('b13-conf-liquid-h').setAttribute('y', 200 - fillConfH);

        const fillConfM = (pctConfM / 100) * 200;
        document.getElementById('b13-conf-liquid-m').setAttribute('height', fillConfM);
        document.getElementById('b13-conf-liquid-m').setAttribute('y', 200 - fillConfM);

    }, 50);

    })
    .catch(error => {
    console.error('Detalle del error en Bloque 13:', error);
    });
});

// =========================================================================
// AUTOMATIZACIÓN BLOQUE 14: INFORMES Y REPORTES INSTITUCIONALES
// =========================================================================
window.addEventListener('DOMContentLoaded', () => {
fetchSheetData('14_Reportes').then(data => {

    let textosInterfase = {};
    let reportes = [];

    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        if (!row || row.length === 0) continue;

        // Mapeo Textos Generales (Ajustado a Columnas I y J -> Índices 8 y 9)
        const keyTexto = (row[8] || '').trim();
        if (keyTexto) {
        textosInterfase[keyTexto] = { j: (row[9] || '').trim() };
        }

        // Mapeo Datos de las Tarjetas (Columnas A - F -> Índices 0 - 5)
        let cat = (row[0] || '').trim();
        let titulo = (row[1] || '').trim();
        let badge = (row[2] || '').trim();
        let desc = (row[3] || '').trim();
        let btn = (row[4] || '').trim();
        let url = (row[5] || '').trim();

        if (cat && cat !== 'Categoría (Tag)') {
        reportes.push({ cat, titulo, badge, desc, btn, url });
        }
    }

    const b14 = document.getElementById('bloque-14-cards');
    if (!b14) return;

    // 1. Inyectar Textos Generales
    if (textosInterfase['PRIMERA LÍNEA']) document.getElementById('b14-tag-text').innerText = textosInterfase['PRIMERA LÍNEA'].j;
    if (textosInterfase['LÍNEA PRINCIPAL']) document.getElementById('b14-title').innerText = textosInterfase['LÍNEA PRINCIPAL'].j;
    if (textosInterfase['INTRODUCCIÓN']) document.getElementById('b14-lead').innerText = textosInterfase['INTRODUCCIÓN'].j;
    if (textosInterfase['FECHA DE CORTE']) document.getElementById('b14-date').innerText = `Corte institucional ${textosInterfase['FECHA DE CORTE'].j}`;

    // 2. Inyectar Datos Tarjeta 1
    if (reportes.length > 0) {
        document.getElementById('b14-c1-tag').innerText = reportes[0].cat;
        document.getElementById('b14-c1-title').innerText = reportes[0].titulo;
        // document.getElementById('b14-c1-badge').innerText = reportes[0].badge;
        document.getElementById('b14-c1-desc').innerText = reportes[0].desc;
        
        const link1 = document.getElementById('b14-c1-link');
        link1.innerHTML = `${reportes[0].btn} &#8599;`;
        link1.href = reportes[0].url;
    }

    // 3. Inyectar Datos Tarjeta 2
    if (reportes.length > 1) {
        document.getElementById('b14-c2-tag').innerText = reportes[1].cat;
        document.getElementById('b14-c2-title').innerText = reportes[1].titulo;
        // document.getElementById('b14-c2-badge').innerText = reportes[1].badge;
        document.getElementById('b14-c2-desc').innerText = reportes[1].desc;
        
        const link2 = document.getElementById('b14-c2-link');
        link2.innerHTML = `${reportes[1].btn} &#8599;`;
        link2.href = reportes[1].url;
    }

    })
    .catch(error => {
    console.error('Detalle del error en Bloque 14:', error);
    });
});

// =========================================================================
// AUTOMATIZACIÓN BLOQUE 15: OTROS DATOS (CORRECCIÓN ESTRICTA DE TÍTULOS)
// =========================================================================
window.addEventListener('DOMContentLoaded', () => {
fetchSheetData('15_Otros').then(data => {
    let datos = {};
    let desgloseFeria = [];

    const parseDecimal = (val) => parseFloat(String(val).replace(/[^\d.-]/g, '')) || 0;

    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        // Validar que la fila exista y tenga al menos hasta la columna C (índice 2)
        if (!row || row.length < 3) continue; 

        // Limpieza estricta: un solo espacio entre palabras, sin acentos, todo mayúsculas
        let colB = String(row[1] || '')
        .trim()
        .toUpperCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .replace(/\s+/g, ' '); 
        
        let colC = String(row[2] || '').trim();
        let colD = String(row[3] || '').trim(); // Columna D
        let colE = String(row[4] || '').trim(); // Columna E (URLs)

        if (colB === 'CONCEPTO' || colB === '') continue;

        // Asignación con coincidencia exacta para evitar falsos positivos
        if (colB === 'PRIMERA LINEA') {
        datos['TAG'] = colC;
        } else if (colB === 'LINEA PRINCIPAL') {
        datos['TITLE'] = colC;
        } else if (colB === 'DESGLOSE') {
        desgloseFeria.push({ material: colC, kilos: parseDecimal(colD), img: colE });
        } else {
        datos[colB] = colC; 
        }
    }

    const b15 = document.getElementById('bloque-15-cards');
    if (!b15) return;

    // 1. Textos Generales 
    document.getElementById('b15-tag-text').innerText = datos['TAG'] || 'OTROS DATOS';
    document.getElementById('b15-title').innerText = datos['TITLE'] || 'Sobre nuestra Universidad';

    // 2. Historias Narrativas 
    const movNac = datos['NACIONAL'] || '0';
    const movInt = datos['INTERNACIONAL'] || '0';
    document.getElementById('b15-mov-desc').innerHTML = `La proyección global de nuestra comunidad estudiantil se fortalece a través de los programas de intercambio. Actualmente contamos con una participación de <strong style="color: var(--uaa-navy); font-size: 30px;">${movNac}</strong> estudiantes en movilidad académica a nivel nacional y <strong style="color: var(--uaa-navy); font-size: 30px;">${movInt}</strong> estudiantes a nivel internacional.`;

    const incTot = datos['TOTAL PROYECTOS'] || '0';
    const incStart = datos['STARTUPS'] || '0';
    document.getElementById('b15-inc-desc').innerHTML = `Fomentamos el emprendimiento universitario impulsando iniciativas innovadoras. Nuestra incubadora respalda activamente <strong style="color: var(--uaa-navy); font-size: 30px;">${incTot}</strong> proyectos de negocios, de los cuales <strong style="color: var(--uaa-gold-vibrant); font-size: 30px;">${incStart}</strong> han logrado consolidarse exitosamente como <em>startups</em> en el mercado actual.`;

    // 3. Párrafos de la Feria Ambiental
    const pFeria = `Durante ${datos['FECHA'] || ''} se llevó a cabo la ${datos['EDICION'] || ''} ${datos['NOMBRE'] || ''}, ${datos['CUERPO'] || ''}`;
    document.getElementById('b15-feria-p1').innerText = pFeria;
    document.getElementById('b15-feria-res').innerText = datos['RESULTADOS'] || '';
    
    const feriaConcl = document.getElementById('b15-feria-concl');
    if (datos['CONCLUSION']) {
        feriaConcl.innerText = datos['CONCLUSION'];
        feriaConcl.style.display = 'block';
    }

    // 4. Inyección de Tarjetas con las Imágenes PNG
    let htmlDesglose = '';
    desgloseFeria.forEach(item => {
        const imgTag = item.img ? `<img src="${item.img}" alt="${item.material}" style="width: 100%; height: 100%; object-fit: contain; opacity: 0.85;">` : '';

        htmlDesglose += `
        <div style="background: #FFFFFF; border-radius: 16px; padding: 20px 24px; display: flex; align-items: center; gap: 18px; box-shadow: 0 4px 12px rgba(22, 48, 114, 0.05); transition: transform 0.2s, box-shadow 0.2s; cursor: default;" onmouseover="this.style.transform='translateY(-3px)'; this.style.boxShadow='0 12px 24px rgba(22, 48, 114, 0.12)'" onmouseout="this.style.transform='none'; this.style.boxShadow='0 4px 12px rgba(22, 48, 114, 0.05)'">
            <div style="width: 52px; height: 52px; border-radius: 12px; background: rgba(21, 96, 130, 0.05); display: flex; align-items: center; justify-content: center; flex-shrink: 0; padding: 10px;">
            <div style="width: 100%; height: 100%;">
                ${imgTag}
            </div>
            </div>
            <div>
            <div style="font-size: 12px; font-weight: 700; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.08em; margin-bottom: 4px;">${item.material}</div>
            <div style="font-size: 24px; font-weight: 800; color: var(--uaa-navy); line-height: 1;">
                <span class="counter-dec" data-target="${item.kilos}">0</span><span style="font-size: 14px; font-weight: 600; color: var(--text-secondary); margin-left: 6px;">kg</span>
            </div>
            </div>
        </div>
        `;
    });
    document.getElementById('b15-feria-desglose').innerHTML = htmlDesglose;

    const animarDatoDecimal = (elemento, objetivo) => {
        const duration = 1400;
        const startTime = performance.now();
        const update = (currentTime) => {
        const progress = Math.min((currentTime - startTime) / duration, 1);
        const easeOut = 1 - Math.pow(1 - progress, 3);
        const actual = easeOut * objetivo;
        elemento.innerText = actual.toLocaleString('es-MX', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        
        if (progress < 1) {
            requestAnimationFrame(update);
        } else {
            elemento.innerText = objetivo.toLocaleString('es-MX', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        }
        };
        requestAnimationFrame(update);
    };

    // 5. Detonador de animaciones
    setTimeout(() => {
        const cardsContainer = document.getElementById('bloque-15-cards');
        if(!cardsContainer) return;

        cardsContainer.querySelectorAll('.counter').forEach(c => {
        if(typeof animarDato === 'function') animarDato(c, parseFloat(c.getAttribute('data-target') || 0));
        });
        cardsContainer.querySelectorAll('.counter-dec').forEach(c => {
        animarDatoDecimal(c, parseFloat(c.getAttribute('data-target') || 0));
        });
    }, 50);

})
.catch(error => console.error('Detalle del error en Bloque 15:', error));
});

// =========================================================================
// AUTOMATIZACIÓN BLOQUE: DUDAS / CONTACTO
// =========================================================================
window.addEventListener('DOMContentLoaded', () => {
fetchSheetData('Dudas').then(data => {
    
    // 1. Objeto para almacenar toda la información temporalmente
    let infoContacto = { correo: '', direccion: '', telefono: '', extension: '', edificio: '' };

    // 2. Leer todas las filas de la base de datos
    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        if (!row || row.length === 0) continue;

        const etiqueta = (row[0] || '').trim();
        const valor = (row[1] || '').trim();

        if (etiqueta.includes('Correo')) infoContacto.correo = valor;
        else if (etiqueta.includes('Direcci')) infoContacto.direccion = valor;
        else if (etiqueta.includes('Teléf')) infoContacto.telefono = valor;
        else if (etiqueta.includes('Extensi')) infoContacto.extension = valor;
        else if (etiqueta.includes('Edificio')) infoContacto.edificio = valor;
    }

    // 3. INYECTAR LA INFORMACIÓN AL HTML
    
    // Inyectar el Correo
    if (infoContacto.correo) {
        const linkCorreo = document.getElementById('contacto-correo-link');
        if (linkCorreo) {
        linkCorreo.innerText = infoContacto.correo;
        linkCorreo.href = `mailto:${infoContacto.correo}`; 
        }
    }

    // Inyectar Dirección
    if (infoContacto.direccion) {
        const elDireccion = document.getElementById('contacto-direccion');
        if (elDireccion) elDireccion.innerText = infoContacto.direccion;
    }

    // Inyectar Teléfono y Extensión (combinados)
    if (infoContacto.telefono || infoContacto.extension) {
        const elTelefono = document.getElementById('contacto-telefono');
        let textoTelefono = infoContacto.telefono;
        // Si hay extensión, la agregamos en un renglón nuevo
        if (infoContacto.extension) {
        textoTelefono += `\n${infoContacto.extension}`;
        }
        if (elTelefono) elTelefono.innerText = textoTelefono;
    }

    // Inyectar Edificio
    if (infoContacto.edificio) {
        const elEdificio = document.getElementById('contacto-edificio');
        if (elEdificio) elEdificio.innerText = infoContacto.edificio;
    }

    })
    .catch(error => {
    console.error('Detalle del error en Bloque Dudas:', error);
    });
});

/* =====================================================
   CHATBOT UAA EN NÚMEROS
   Control de apertura y cierre
   ===================================================== */

const chatbotLauncher = document.getElementById('chatbot-launcher');
const chatbotWindow = document.getElementById('chatbot-window');
const chatbotClose = document.getElementById('chatbot-close');

if (chatbotLauncher && chatbotWindow && chatbotClose) {

    // Abrir chatbot con animación
    chatbotLauncher.addEventListener('click', () => {

        chatbotWindow.style.display = 'flex';
        chatbotLauncher.style.display = 'none';

        // Forzar reflow para que la transición CSS se active
        void chatbotWindow.offsetHeight;
        chatbotWindow.classList.add('chatbot-open');

    });

    // Cerrar chatbot con animación
    chatbotClose.addEventListener('click', () => {

        chatbotWindow.classList.remove('chatbot-open');

        // Esperar a que termine la transición antes de ocultar
        chatbotWindow.addEventListener('transitionend', function handler() {
            chatbotWindow.removeEventListener('transitionend', handler);
            chatbotWindow.style.display = 'none';
            chatbotLauncher.style.display = 'flex';
        });

    });

}

/* =====================================================
   Envío de mensajes
   ===================================================== */

const chatbotInput = document.getElementById('chatbot-input');
const chatbotSend = document.getElementById('chatbot-send');
const chatbotMessages = document.getElementById('chatbot-messages');

async function enviarMensajeChatbot() {

    const pregunta = chatbotInput.value.trim();

    // No enviar mensajes vacíos
    if (!pregunta) {
        return;
    }

    // Mostrar pregunta del usuario
    const mensajeUsuario = document.createElement('div');

    mensajeUsuario.className = 'chatbot-message user-message';

    mensajeUsuario.innerHTML = `
        <div class="message-bubble">
            ${pregunta}
        </div>
    `;

    chatbotMessages.appendChild(mensajeUsuario);

    // Limpiar caja de texto
    chatbotInput.value = '';

    // Mostrar indicador de carga
    const chatbotLoading = document.getElementById('chatbot-loading');

    if (chatbotLoading) {
        chatbotLoading.style.display = 'flex';
    }

    // Desplazar conversación hacia abajo
    chatbotMessages.scrollTop = chatbotMessages.scrollHeight;

    try {

        console.log('📩 Enviando pregunta a Render:', pregunta);

        const response = await fetch(
            'https://chatbot-uaa-en-numeros.onrender.com/api/chat',
            {
                method: 'POST',

                headers: {
                    'Content-Type': 'application/json'
                },

                body: JSON.stringify({
                    pregunta: pregunta
                })
            }
        );

        const resultado = await response.json();

        console.log('📥 Respuesta de Render:', resultado);

        if (!response.ok) {
            throw new Error(
                resultado.error || 'Error al comunicarse con el servidor.'
            );
        }

        // Crear mensaje del bot
        const mensajeBot = document.createElement('div');

        mensajeBot.className = 'chatbot-message bot-message';

        mensajeBot.innerHTML = `
            <div class="message-bubble">
                ${resultado.respuesta || 'No se recibió una respuesta.'}
            </div>
        `;

        chatbotMessages.appendChild(mensajeBot);

    } catch (error) {

        console.error('❌ Error del chatbot:', error);

        const mensajeError = document.createElement('div');

        mensajeError.className = 'chatbot-message bot-message';

        mensajeError.innerHTML = `
            <div class="message-bubble">
                Lo siento, ocurrió un problema al consultar la información. 
                Por favor, intenta nuevamente.
            </div>
        `;

        chatbotMessages.appendChild(mensajeError);

    } finally {

        // Ocultar indicador de carga
        if (chatbotLoading) {
            chatbotLoading.style.display = 'none';
        }

        // Llevar conversación al final
        chatbotMessages.scrollTop = chatbotMessages.scrollHeight;

    }
}


/* Botón enviar */

if (chatbotSend) {

    chatbotSend.addEventListener('click', () => {

        enviarMensajeChatbot();

    });

}


/* Tecla Enter */

if (chatbotInput) {

    chatbotInput.addEventListener('keydown', (event) => {

        if (event.key === 'Enter') {

            event.preventDefault();

            enviarMensajeChatbot();

        }

    });

}