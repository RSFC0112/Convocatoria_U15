const { createClient } = require('@supabase/supabase-js');
const XLSX = require('xlsx');
const axios = require('axios');
const fs = require('fs');
const path = require('path');

const SUPABASE_URL = 'https://fjbbrzhqlvbkliskslqs.supabase.co';
const SUPABASE_KEY = 'sb_publishable_QRpv4Rs_fIJ8kXvTegR25w_t2SRSExp';
const supabase = createClient(SUPABASE_URL, SUPABASE_KEY);

function normalizarTexto(valor) {
    return String(valor ?? '')
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, ' ')
        .trim();
}

function cargarJugadoresU15DelIndex() {
    const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
    const bloqueJugadores = /const baseDeDatosJugadores = \[([\s\S]*?)\];/.exec(html)?.[1];
    if (!bloqueJugadores) throw new Error('No se encontró baseDeDatosJugadores en index.html.');

    const patronJugador = /\{\s*id:\s*\d+,\s*dorsal:\s*[^,]+,\s*nombre:\s*"([^"]+)",[^}]*?\bposicion:\s*"([^"]+)"/g;
    const jugadores = Array.from(bloqueJugadores.matchAll(patronJugador), coincidencia => ({
        nombre: normalizarTexto(coincidencia[1]),
        posicion: coincidencia[2]
    }));
    if (jugadores.length === 0) throw new Error('No se encontraron jugadores con posición en index.html.');
    return jugadores;
}

function obtenerPosicionU15(nombre, apellido1, jugadoresIndex) {
    const primerNombre = normalizarTexto(nombre).split(' ')[0];
    const primerApellido = normalizarTexto(apellido1);
    const coincidencias = jugadoresIndex.filter(jugador => {
        const tokens = jugador.nombre.split(' ');
        return tokens.includes(primerNombre) && tokens.includes(primerApellido);
    });

    if (coincidencias.length !== 1) {
        console.warn(`  -> No se pudo asociar una posición única a ${nombre} ${apellido1 || ''}.`);
        return null;
    }
    return coincidencias[0].posicion;
}

function obtenerColumnasDatosJugador(filas) {
    const encabezados = filas.find(fila =>
        fila?.some(celda => normalizarTexto(celda) === 'dorsal') &&
        fila.some(celda => normalizarTexto(celda) === 'fecha nac')
    );
    if (!encabezados) return null;

    return {
        dorsal: encabezados.findIndex(celda => normalizarTexto(celda) === 'dorsal'),
        fechaNacimiento: encabezados.findIndex(celda => normalizarTexto(celda) === 'fecha nac')
    };
}

function parsearFechaExcel(valorExcel) {
    if (valorExcel === undefined || valorExcel === null || valorExcel === '') return null;
    if (valorExcel instanceof Date) return Number.isNaN(valorExcel.getTime()) ? null : valorExcel;

    const num = Number(valorExcel);
    if (Number.isFinite(num) && num > 1000) {
        const utcDays = Math.floor(num - 25569);
        const fechaObj = new Date(utcDays * 86400 * 1000);
        return Number.isNaN(fechaObj.getTime()) ? null : fechaObj;
    }

    const texto = String(valorExcel).trim();
    let partes = texto.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (partes) {
        return new Date(Date.UTC(Number(partes[1]), Number(partes[2]) - 1, Number(partes[3])));
    }

    partes = texto.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (partes) {
        const dia = Number(partes[1]);
        const mes = Number(partes[2]);
        const anio = Number(partes[3]);
        const fecha = new Date(Date.UTC(anio, mes - 1, dia));
        if (fecha.getUTCMonth() !== mes - 1 || fecha.getUTCDate() !== dia) return null;
        return fecha;
    }

    const fechaParsed = new Date(texto);
    return Number.isNaN(fechaParsed.getTime()) ? null : fechaParsed;
}

function obtenerEstadoEntrenamiento(valor) {
    if (valor === undefined || valor === null || String(valor).trim() === '') return null;
    const numero = Number(valor);
    return [0, 0.5, 1].includes(numero) ? numero : null;
}

function obtenerEstadoPartido(valor) {
    const texto = String(valor ?? '').trim().toUpperCase().replace(/\s+/g, '');
    if (texto === 'X' || texto === 'P(X)') return 0;
    if (texto === '1' || texto === '1.0') return 1;
    return null;
}

function obtenerNumeroResumen(valor) {
    const numero = Number(valor);
    return valor !== undefined && valor !== null && valor !== '' && Number.isFinite(numero) ? numero : 0;
}

function formatearFechaDDMMYYYY(valorExcel) {
    if (valorExcel === undefined || valorExcel === null || valorExcel === '') return 'NA';
    const fechaObj = parsearFechaExcel(valorExcel);
    if (!fechaObj) return String(valorExcel).trim();
    const anio = fechaObj.getUTCFullYear();
    const mes = String(fechaObj.getUTCMonth() + 1).padStart(2, '0');
    const dia = String(fechaObj.getUTCDate()).padStart(2, '0');
    return `${dia}/${mes}/${anio}`;
}

async function sincronizarU15() {
    try {
        const urlExcelEnLinea = 'https://docs.google.com/spreadsheets/d/1gpY4TcpxmBebSk9Popp5IK7tYVriTTT-/export?format=xlsx';

        console.log('Descargando datos en tiempo real desde Google Sheets...');
        const respuesta = await axios.get(urlExcelEnLinea, { responseType: 'arraybuffer' });
        const workbook = XLSX.read(respuesta.data, { type: 'buffer' });

        const nombreHoja = 'ACTIVOS 2026';
        if (!workbook.SheetNames.includes(nombreHoja)) {
            console.error(`Error: No se encontró la pestaña "${nombreHoja}".`);
            return;
        }
        console.log(`Leyendo la pestaña en línea: "${nombreHoja}"`);

        const worksheet = workbook.Sheets[nombreHoja];
        const filas = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        if (filas.length === 0) {
            console.log("La hoja de cálculo está vacía.");
            return;
        }

        const columnasDatosJugador = obtenerColumnasDatosJugador(filas);
        if (!columnasDatosJugador || columnasDatosJugador.dorsal < 0 || columnasDatosJugador.fechaNacimiento < 0) {
            console.error('Error: No se encontraron las columnas DORSAL y FECHA NAC. en el Excel.');
            return;
        }

        const jugadoresU15DelIndex = cargarJugadoresU15DelIndex();
        const filaFechas = filas[6] || [];
        const filaSecciones = filas[3] || [];
        const indiceInicioEntrenamientos = filaSecciones.findIndex(celda => normalizarTexto(celda).includes('asistencia'));
        const indiceInicioConvocatorias = filaSecciones.findIndex(celda => normalizarTexto(celda).includes('convocatorias'));

        if (indiceInicioEntrenamientos < 0 || indiceInicioConvocatorias <= indiceInicioEntrenamientos) {
            throw new Error('No se encontraron los encabezados de ASISTENCIA y CONVOCATORIAS en el Excel.');
        }

        const indicesU15 = [];
        for (let i = 0; i < filas.length; i++) {
            const fila = filas[i];
            if (!fila) continue;
            const nombre = String(fila[5] || '').trim();
            const esU15 = fila.some(celda =>
                /\bU\s*-?\s*15\b/i.test(String(celda || '').normalize('NFKC'))
            );
            if (esU15 && nombre !== '') {
                indicesU15.push(i);
            }
        }

        console.log(`Jugadores U15 detectados: ${indicesU15.length}`);

        if (indicesU15.length === 0) {
            console.error("Error: No se encontraron jugadores U15 en la hoja en línea.");
            return;
        }

        const entrenamientosOficiales = [];
        for (let idx = indiceInicioEntrenamientos; idx < indiceInicioConvocatorias; idx++) {
            const celdaFecha = filaFechas[idx];
            if (celdaFecha === undefined || celdaFecha === null || String(celdaFecha).trim() === '') continue;

            let huboEntrenamiento = false;
            for (const rowIdx of indicesU15) {
                if (obtenerEstadoEntrenamiento(filas[rowIdx][idx]) !== null) {
                    huboEntrenamiento = true;
                    break;
                }
            }

            if (huboEntrenamiento) {
                const fechaObj = parsearFechaExcel(celdaFecha);
                if (!fechaObj) continue;
                entrenamientosOficiales.push({ colIndex: idx, fecha: fechaObj });
            }
        }

        const partidosOficiales = [];
        for (let idx = indiceInicioConvocatorias; idx < filaFechas.length; idx++) {
            const celdaFecha = filaFechas[idx];
            if (celdaFecha === undefined || celdaFecha === null || String(celdaFecha).trim() === '') continue;

            const huboPartido = indicesU15.some(rowIdx => obtenerEstadoPartido(filas[rowIdx][idx]) !== null);
            if (!huboPartido) continue;

            const fecha = parsearFechaExcel(celdaFecha);
            if (fecha) partidosOficiales.push({ colIndex: idx, fecha });
        }

        let contadorU15 = 0;

        for (const rowIdx of indicesU15) {
            const fila = filas[rowIdx];
            const nombre = fila[5];
            const apellido1 = fila[3];

            if (!nombre) continue;

            contadorU15++;
            const nombreCompleto = `${String(nombre).trim()} ${String(apellido1 || '').trim()}`;
            const fechaInscripcion = parsearFechaExcel(fila[15]);
            const anioInscripcion = fechaInscripcion?.getUTCFullYear();
            const inscripcionDesde2026 = anioInscripcion >= 2026;
            const cuentaDesdeInscripcion = item => !inscripcionDesde2026
                || !fechaInscripcion
                || item.fecha >= fechaInscripcion;
            const entrenamientosContables = entrenamientosOficiales.filter(cuentaDesdeInscripcion);
            const partidosContables = partidosOficiales.filter(cuentaDesdeInscripcion);
            const entrenamientosDelJugador = entrenamientosContables.filter(item =>
                obtenerEstadoEntrenamiento(fila[item.colIndex]) !== null
            );
            const partidosDelJugador = partidosContables.filter(item =>
                obtenerEstadoPartido(fila[item.colIndex]) !== null
            );

            const cumplidos = entrenamientosDelJugador.reduce((total, item) =>
                total + (obtenerEstadoEntrenamiento(fila[item.colIndex]) === 1 ? 1 : 0), 0
            );
            const jugadosDesdeCeldas = partidosDelJugador.reduce((total, item) =>
                total + (obtenerEstadoPartido(fila[item.colIndex]) === 1 ? 1 : 0), 0
            );
            const ofrecidos = inscripcionDesde2026
                ? entrenamientosContables.length
                : obtenerNumeroResumen(fila[22]);
            const jugados = inscripcionDesde2026
                ? jugadosDesdeCeldas
                : obtenerNumeroResumen(fila[145]);
            const programados = inscripcionDesde2026
                ? partidosContables.length
                : obtenerNumeroResumen(fila[146]);
            const textoEntrenamientos = `${cumplidos}/${ofrecidos}`;
            const textoPartidos = `${jugados}/${programados}`;

            const sesionesUltimosCinco = entrenamientosDelJugador.length > 5
                ? entrenamientosDelJugador.slice(-5)
                : entrenamientosOficiales.slice(-5);
            const ultimasFechasData = sesionesUltimosCinco.map(item => {
                const dia = String(item.fecha.getUTCDate()).padStart(2, '0');
                const mes = String(item.fecha.getUTCMonth() + 1).padStart(2, '0');
                const anio = item.fecha.getUTCFullYear();
                const estado = obtenerEstadoEntrenamiento(fila[item.colIndex]);
                const noInscritoAun = inscripcionDesde2026 && fechaInscripcion && item.fecha < fechaInscripcion;
                const textoMostrar = noInscritoAun ? 'NA' : estado === 1 ? '1' : '0';

                return { fecha: `${anio}-${mes}-${dia}`, asistio: textoMostrar };
            });

            const fechaInscripcionLimpia = formatearFechaDDMMYYYY(fila[15]);
            const dorsalExcel = fila[columnasDatosJugador.dorsal];
            const textoDorsal = String(dorsalExcel ?? '').trim();
            const dorsal = textoDorsal === '' || textoDorsal.toUpperCase() === 'NULL' ? '-' : textoDorsal;
            const fechaNacimiento = parsearFechaExcel(fila[columnasDatosJugador.fechaNacimiento]);
            const posicion = obtenerPosicionU15(nombre, apellido1, jugadoresU15DelIndex);

            const datosJugador = {
                id_jugador: Number(fila[1]) || contadorU15,
                nombre: nombreCompleto,
                dorsal,
                posicion,
                anio_nacimiento: fechaNacimiento?.getUTCFullYear() ?? null,
                entrenamientos: textoEntrenamientos,
                partidos: textoPartidos,
                fecha_inscripcion: fechaInscripcionLimpia,
                ultimas_fechas: ultimasFechasData
            };

            const { error } = await supabase.from('Estadísticas_U15').upsert(datosJugador, { onConflict: 'id_jugador' });

            if (error) {
                console.error(`  -> Error al subir a ${datosJugador.nombre}:`, error.message);
            } else {
                console.log(`  -> Sincronizado U15: ${datosJugador.nombre} | Entrenamientos: ${textoEntrenamientos} | Partidos: ${textoPartidos}`);
            }
        }

        console.log(`¡Sincronización completa! Se procesaron ${contadorU15} jugadores de la U15 en línea.`);

    } catch (error) {
        console.error('Ocurrió un error general descargando de Google Sheets:', error.message);
    }
}

sincronizarU15();