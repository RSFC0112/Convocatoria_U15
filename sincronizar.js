const { createClient } = require('@supabase/supabase-js');
const XLSX = require('xlsx');
const axios = require('axios');

const SUPABASE_URL = 'https://fjbbrzhqlvbkliskslqs.supabase.co';
const SUPABASE_KEY = 'sb_publishable_QRpv4Rs_fIJ8kXvTegR25w_t2SRSExp';
const supabase = createClient(SUPABASE_URL, SUPABASE_KEY);
const ANIO_ACTUAL = new Date().getUTCFullYear();
const normalizarNombre = nombre => String(nombre || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
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
        const anio = Number(partes[1]);
        const mes = Number(partes[2]);
        const dia = Number(partes[3]);
        const fecha = new Date(Date.UTC(anio, mes - 1, dia));
        if (fecha.getUTCFullYear() !== anio || fecha.getUTCMonth() !== mes - 1 || fecha.getUTCDate() !== dia) return null;
        return fecha;
    }

    partes = texto.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (partes) {
        const dia = Number(partes[1]);
        const mes = Number(partes[2]);
        const anio = Number(partes[3]);
        const fecha = new Date(Date.UTC(anio, mes - 1, dia));
        if (fecha.getUTCFullYear() !== anio || fecha.getUTCMonth() !== mes - 1 || fecha.getUTCDate() !== dia) return null;
        return fecha;
    }

    return null;
}

function formatearFechaDDMMYYYY(valorExcel) {
    if (valorExcel === undefined || valorExcel === null || valorExcel === '') return 'NA';
    const fechaObj = parsearFechaExcel(valorExcel);
    if (!fechaObj) return 'NA';
    const anio = fechaObj.getUTCFullYear();
    const mes = String(fechaObj.getUTCMonth() + 1).padStart(2, '0');
    const dia = String(fechaObj.getUTCDate()).padStart(2, '0');
    return `${dia}/${mes}/${anio}`;
}

function obtenerDiaUTC(fecha) {
    return Date.UTC(fecha.getUTCFullYear(), fecha.getUTCMonth(), fecha.getUTCDate());
}

function esFechaIgualOPosterior(fecha, fechaLimite) {
    return obtenerDiaUTC(fecha) >= obtenerDiaUTC(fechaLimite);
}

async function sincronizarCategorias() {
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

        const filaFechas = filas[6] || [];
        const normalizarEncabezado = valor => String(valor || '')
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .toLowerCase()
            .replace(/[^a-z0-9]+/g, ' ')
            .trim();
        let indiceNombre = -1;
        let indiceApellido1 = -1;
        let indiceDorsal = -1;
        let indiceFechaNacimiento = -1;
        let indiceFechaInscripcion = -1;
        let indiceEntrenamientosCumplidos = -1;
        let indiceEntrenamientosOfrecidos = -1;
        let indicePartidosJugados = -1;
        let indicePartidosProgramados = -1;

        for (const fila of filas.slice(0, 10)) {
            (fila || []).forEach((valor, indice) => {
                const encabezado = normalizarEncabezado(valor);
                if (/\bnombres?\b/.test(encabezado) && indiceNombre === -1) {
                    indiceNombre = indice;
                }
                if ((/\bapellido\s*(1|uno)\b/.test(encabezado)
                    || /\bprimer apellido\b/.test(encabezado)
                    || /\bapellido paterno\b/.test(encabezado))
                    && indiceApellido1 === -1) {
                    indiceApellido1 = indice;
                }
                if (encabezado === 'dorsal' && indiceDorsal === -1) {
                    indiceDorsal = indice;
                }
                if (/^fecha nac\b/.test(encabezado) && indiceFechaNacimiento === -1) {
                    indiceFechaNacimiento = indice;
                }
                if ((/\bfecha\b.*\b(inscripcion|ingreso|registro)\b/.test(encabezado)
                    || /\b(inscripcion|ingreso|registro)\b.*\bfecha\b/.test(encabezado))
                    && indiceFechaInscripcion === -1) {
                    indiceFechaInscripcion = indice;
                }
                if (encabezado === 'anual cumplido' && indiceEntrenamientosCumplidos === -1) {
                    indiceEntrenamientosCumplidos = indice;
                }
                if (encabezado === 'anual ofrecido' && indiceEntrenamientosOfrecidos === -1) {
                    indiceEntrenamientosOfrecidos = indice;
                }
                if (encabezado === 'jugados' && indicePartidosJugados === -1) {
                    indicePartidosJugados = indice;
                }
                if (encabezado === 'programados' && indicePartidosProgramados === -1) {
                    indicePartidosProgramados = indice;
                }
            });
            if (indiceNombre !== -1 && indiceApellido1 !== -1 && indiceDorsal !== -1
                && indiceFechaNacimiento !== -1 && indiceFechaInscripcion !== -1
                && indiceEntrenamientosCumplidos !== -1 && indiceEntrenamientosOfrecidos !== -1
                && indicePartidosJugados !== -1 && indicePartidosProgramados !== -1) break;
        }

        if ([indiceNombre, indiceApellido1, indiceDorsal, indiceFechaNacimiento, indiceFechaInscripcion, indiceEntrenamientosCumplidos,
            indiceEntrenamientosOfrecidos, indicePartidosJugados, indicePartidosProgramados].includes(-1)) {
            console.error('Error: Faltan encabezados de identidad, dorsal, fecha de nacimiento, inscripción o estadísticas anuales en las primeras filas.');
            return;
        }
        console.log(`Columnas detectadas: nombre ${indiceNombre + 1}, primer apellido ${indiceApellido1 + 1}, dorsal ${indiceDorsal + 1}, fecha de nacimiento ${indiceFechaNacimiento + 1}, fecha de inscripción ${indiceFechaInscripcion + 1}, entrenamientos ${indiceEntrenamientosCumplidos + 1}/${indiceEntrenamientosOfrecidos + 1}, partidos ${indicePartidosJugados + 1}/${indicePartidosProgramados + 1}`);

        const categorias = ['U13', 'U15'];
        const jugadoresPorCategoria = new Map();
        const siguienteIdPorCategoria = new Map();
        for (const categoria of categorias) {
            const { data, error } = await supabase
                .from(`Estadísticas_${categoria}`)
                .select('id_jugador,nombre,posicion');

            if (error) {
                console.error(`Error al consultar la plantilla ${categoria}:`, error.message);
                return;
            }

            const jugadores = new Map((data || [])
                .filter(jugador => jugador.nombre && jugador.id_jugador !== null)
                .map(jugador => [normalizarNombre(jugador.nombre), jugador]));
            const ids = Array.from(jugadores.values(), jugador => Number(jugador.id_jugador))
                .filter(Number.isSafeInteger);
            jugadoresPorCategoria.set(categoria, jugadores);
            siguienteIdPorCategoria.set(categoria, ids.length > 0 ? Math.max(...ids) + 1 : 1);
        }

        const indicesJugadores = [];
        for (let i = 0; i < filas.length; i++) {
            const fila = filas[i];
            if (!fila) continue;
            const nombre = String(fila[indiceNombre] || '').trim();
            if (nombre === '') continue;

            const categoria = categorias.find(item => fila.some(celda =>
                new RegExp(`\\bU\\s*-?\\s*${item.slice(1)}\\b`, 'i')
                    .test(String(celda || '').normalize('NFKC'))
            ));
            if (!categoria) continue;

            const nombreCompleto = `${nombre} ${String(fila[indiceApellido1] || '').trim()}`.trim();
            const jugadores = jugadoresPorCategoria.get(categoria);
            let jugadorPlantilla = jugadores.get(normalizarNombre(nombreCompleto));
            if (!jugadorPlantilla) {
                jugadorPlantilla = {
                    id_jugador: siguienteIdPorCategoria.get(categoria),
                    nombre: nombreCompleto,
                    posicion: 'INV'
                };
                siguienteIdPorCategoria.set(categoria, jugadorPlantilla.id_jugador + 1);
                jugadores.set(normalizarNombre(nombreCompleto), jugadorPlantilla);
            }
            indicesJugadores.push({ rowIdx: i, categoria, jugadorPlantilla });
        }

        if (indicesJugadores.length === 0) {
            console.error('Error: No se encontraron jugadores U13 ni U15 en la hoja en línea.');
            return;
        }

        for (const categoria of categorias) {
            const cantidad = indicesJugadores.filter(jugador => jugador.categoria === categoria).length;
            console.log(`Jugadores ${categoria} detectados: ${cantidad}`);
        }

        const encabezadosSeccion = filas[3] || [];
        const indiceAsistencia = encabezadosSeccion.findIndex(valor =>
            normalizarEncabezado(valor) === 'asistencia 2026'
        );
        const indiceConvocatorias = encabezadosSeccion.findIndex(valor =>
            normalizarEncabezado(valor) === 'convocatorias 2026'
        );
        if (indiceAsistencia === -1 || indiceConvocatorias <= indiceAsistencia) {
            console.error('Error: No se encontraron los encabezados de asistencia y convocatorias.');
            return;
        }

        const columnasEntrenamientos = [];
        for (let idx = indiceAsistencia + 1; idx < indiceConvocatorias; idx++) {
            const fecha = parsearFechaExcel(filaFechas[idx]);
            if (fecha) columnasEntrenamientos.push({ colIndex: idx, fecha });
        }
        columnasEntrenamientos.sort((a, b) => a.fecha - b.fecha);

        const columnasPartidos = [];
        const inicioPartidos = indiceConvocatorias + 1;
        for (let idx = inicioPartidos; idx < indicePartidosJugados; idx++) {
            let fecha = parsearFechaExcel(filaFechas[idx]);
            for (let distancia = 1; !fecha && (idx - distancia >= inicioPartidos || idx + distancia < indicePartidosJugados); distancia++) {
                fecha = idx + distancia < indicePartidosJugados
                    ? parsearFechaExcel(filaFechas[idx + distancia])
                    : null;
                if (!fecha && idx - distancia >= inicioPartidos) {
                    fecha = parsearFechaExcel(filaFechas[idx - distancia]);
                }
            }
            if (fecha) columnasPartidos.push({ colIndex: idx, fecha });
        }
        columnasPartidos.sort((a, b) => a.fecha - b.fecha);

        const obtenerMarcaEntrenamiento = valor => {
            if (valor === undefined || valor === null || String(valor).trim() === '') return null;
            const numero = Number(valor);
            return [0, 0.5, 1].includes(numero) ? numero : null;
        };
        const obtenerMarcaPartido = valor => {
            const marca = String(valor ?? '').trim().toUpperCase();
            return ['X', 'P(X)', '1', '1.0'].includes(marca) ? marca : null;
        };

        const contadores = new Map(categorias.map(categoria => [categoria, 0]));

        for (const { rowIdx, categoria, jugadorPlantilla } of indicesJugadores) {
            const fila = filas[rowIdx];
            const nombre = fila[indiceNombre];
            const apellido1 = fila[indiceApellido1];

            if (!nombre) continue;

            contadores.set(categoria, contadores.get(categoria) + 1);
            const nombreCompleto = `${String(nombre).trim()} ${String(apellido1 || '').trim()}`;
            const fechaInscripcion = parsearFechaExcel(fila[indiceFechaInscripcion]);
            const calcularDesdeInscripcion = fechaInscripcion?.getUTCFullYear() === ANIO_ACTUAL;
            const fechaPermitida = fecha => !calcularDesdeInscripcion || esFechaIgualOPosterior(fecha, fechaInscripcion);
            const obtenerValorAnual = indice => {
                const valor = fila[indice];
                return valor !== undefined && valor !== null && String(valor).trim() !== '' ? valor : 0;
            };
            const entrenamientosDelJugador = columnasEntrenamientos.filter(item =>
                fechaPermitida(item.fecha)
                && obtenerMarcaEntrenamiento(fila[item.colIndex]) !== null
            );
            const entrenamientosAsistidos = obtenerValorAnual(indiceEntrenamientosCumplidos);
            const totalEntrenamientos = calcularDesdeInscripcion
                ? entrenamientosDelJugador.length
                : obtenerValorAnual(indiceEntrenamientosOfrecidos);

            const partidosDelJugador = columnasPartidos.filter(item =>
                fechaPermitida(item.fecha)
                && obtenerMarcaPartido(fila[item.colIndex]) !== null
            );
            const partidosAsistidos = obtenerValorAnual(indicePartidosJugados);
            const totalPartidos = calcularDesdeInscripcion
                ? partidosDelJugador.length
                : obtenerValorAnual(indicePartidosProgramados);

            const textoEntrenamientos = `${entrenamientosAsistidos}/${totalEntrenamientos}`;
            const textoPartidos = `${partidosAsistidos}/${totalPartidos}`;

            const entrenamientosParaUltimasFechas = entrenamientosDelJugador.slice(-5);
            const ultimasFechasData = entrenamientosParaUltimasFechas.map(item => {
                const dia = String(item.fecha.getUTCDate()).padStart(2, '0');
                const mes = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'][item.fecha.getUTCMonth()];
                const marca = obtenerMarcaEntrenamiento(fila[item.colIndex]);

                return {
                    fecha: `${dia}-${mes}`,
                    asistio: marca
                };
            });
            while (ultimasFechasData.length < 5) {
                ultimasFechasData.unshift({ fecha: 'NA', asistio: 'NA' });
            }

            const fechaInscripcionLimpia = formatearFechaDDMMYYYY(fila[indiceFechaInscripcion]);
            const valorDorsal = fila[indiceDorsal];
            const dorsal = valorDorsal !== undefined && valorDorsal !== null && String(valorDorsal).trim() !== ''
                && Number.isFinite(Number(valorDorsal)) ? Number(valorDorsal) : null;
            const fechaNacimiento = parsearFechaExcel(fila[indiceFechaNacimiento]);

            const datosJugador = {
                id_jugador: jugadorPlantilla.id_jugador,
                nombre: nombreCompleto,
                entrenamientos: textoEntrenamientos,
                partidos: textoPartidos,
                fecha_inscripcion: fechaInscripcionLimpia,
                ultimas_fechas: ultimasFechasData,
                dorsal,
                posicion: jugadorPlantilla.posicion,
                anio_nacimiento: fechaNacimiento ? fechaNacimiento.getUTCFullYear() : null
            };

            const { error } = await supabase
                .from(`Estadísticas_${categoria}`)
                .upsert(datosJugador, { onConflict: 'id_jugador' });

            if (error) {
                console.error(`  -> Error al subir a ${datosJugador.nombre}:`, error.message);
            } else {
                console.log(`  -> Sincronizado ${categoria}: ${datosJugador.nombre} | Posición: ${datosJugador.posicion} | Entrenamientos: ${textoEntrenamientos} | Partidos: ${textoPartidos}`);
            }
        }

        for (const categoria of categorias) {
            console.log(`Sincronización ${categoria} completa: ${contadores.get(categoria)} jugadores procesados.`);
        }

    } catch (error) {
        console.error('Ocurrió un error general descargando de Google Sheets:', error.message);
    }
}

sincronizarCategorias();
