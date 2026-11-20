/**
 * TRAMPIL - Registro de clientes, pedidos y pagos
 *
 * Cómo usarlo:
 *   1. Entra a https://script.google.com con tu cuenta de Google -> "Nuevo proyecto".
 *   2. Borra lo que aparece, pega TODO este archivo y guarda (Ctrl + S).
 *   3. Arriba, elige la función "crearFormularioTrampil" y pulsa "Ejecutar".
 *   4. Acepta los permisos (es tu propia cuenta).
 *   5. Abre "Registro de ejecución": ahí salen los links del formulario y del Excel.
 *
 * Qué crea:
 *   - Un Google Form con secciones (cliente, tipo de estudiante, pedido, pago, control).
 *   - Una hoja de cálculo de Google conectada: cada respuesta se llena sola en una fila.
 *     (Se descarga como Excel desde Archivo > Descargar > Microsoft Excel).
 *   - Si subes a tu Google Drive la imagen "Trampil encabezado.png" (logo + personaje),
 *     la pone al inicio del formulario.
 */

var NOMBRE_MARCA = 'Trampil';

function crearFormularioTrampil() {
  var form = FormApp.create(NOMBRE_MARCA + ' - Registro de clientes y pedidos');
  form.setDescription(
    'Apoyo a estudiantes de Arquitectura (del 1.er al 10.º ciclo) y de Estudios Generales (hasta el 5.º ciclo).\n' +
    'Registro interno de clientes, pedidos y pagos (Yape / Plin / transferencia).'
  );
  form.setCollectEmail(false);
  form.setProgressBar(true);
  form.setConfirmationMessage('¡Registro guardado! Gracias por confiar en ' + NOMBRE_MARCA + '.');

  // ---------- Logo y personaje (si están en tu Drive) ----------
  agregarImagenesDeMarca_(form);

  // ---------- SECCIÓN 1: Datos del cliente ----------
  form.addSectionHeaderItem().setTitle('Datos del cliente');

  form.addTextItem().setTitle('Nombre completo').setRequired(true);

  form.addTextItem()
    .setTitle('Número de WhatsApp')
    .setHelpText('9 dígitos, sin espacios. Ej.: 987654321')
    .setValidation(FormApp.createTextValidation()
      .requireTextMatchesPattern('^9[0-9]{8}$')
      .setHelpText('Escribe un celular de 9 dígitos que empiece con 9.')
      .build())
    .setRequired(true);

  form.addTextItem()
    .setTitle('Correo electrónico (opcional)')
    .setValidation(FormApp.createTextValidation().requireTextIsEmail().build());

  form.addTextItem().setTitle('Universidad o instituto').setRequired(true);

  form.addMultipleChoiceItem()
    .setTitle('¿Es cliente nuevo o ya nos compró antes?')
    .setChoiceValues(['Cliente nuevo', 'Cliente recurrente'])
    .setRequired(true);

  form.addMultipleChoiceItem()
    .setTitle('¿Cómo nos conoció?')
    .setChoiceValues(['WhatsApp', 'TikTok', 'Facebook', 'Instagram', 'Recomendación de un amigo'])
    .showOtherOption(true);

  var tipoEstudiante = form.addMultipleChoiceItem()
    .setTitle('Tipo de estudiante')
    .setRequired(true);

  // ---------- SECCIÓN 2A: Arquitectura ----------
  var secArq = form.addPageBreakItem()
    .setTitle('Estudiante de Arquitectura')
    .setHelpText('Del 1.er al 10.º ciclo');

  form.addListItem()
    .setTitle('Ciclo (Arquitectura)')
    .setChoiceValues(['1.er ciclo', '2.º ciclo', '3.er ciclo', '4.º ciclo', '5.º ciclo',
                      '6.º ciclo', '7.º ciclo', '8.º ciclo', '9.º ciclo', '10.º ciclo'])
    .setRequired(true);

  form.addTextItem()
    .setTitle('Curso (Arquitectura)')
    .setHelpText('Ej.: Taller de Arquitectura VI, Instalaciones Eléctricas, Construcción II')
    .setRequired(true);

  // ---------- SECCIÓN 2B: Estudios Generales ----------
  var secGen = form.addPageBreakItem()
    .setTitle('Estudiante de Estudios Generales')
    .setHelpText('Hasta el 5.º ciclo');

  form.addListItem()
    .setTitle('Ciclo (Estudios Generales)')
    .setChoiceValues(['1.er ciclo', '2.º ciclo', '3.er ciclo', '4.º ciclo', '5.º ciclo'])
    .setRequired(true);

  form.addTextItem().setTitle('Carrera').setRequired(true);

  form.addTextItem()
    .setTitle('Curso (Estudios Generales)')
    .setHelpText('Ej.: Matemática Básica, Comunicación, Física')
    .setRequired(true);

  // ---------- SECCIÓN 3: Pedido ----------
  var secPedido = form.addPageBreakItem().setTitle('Detalle del pedido');

  form.addCheckboxItem()
    .setTitle('Tipo de trabajo')
    .setChoiceValues([
      'Láminas / planos (plantas, cortes, elevaciones)',
      'Modelado 3D / renders',
      'Maqueta',
      'Detalles constructivos / mobiliario',
      'Instalaciones eléctricas o sanitarias',
      'Metrados / presupuesto',
      'Expediente técnico',
      'Compatibilización de especialidades',
      'Láminas de presentación / diapositivas',
      'Informe / monografía',
      'Tarea, práctica o examen de curso general'
    ])
    .showOtherOption(true)
    .setRequired(true);

  form.addParagraphTextItem()
    .setTitle('Descripción del pedido')
    .setHelpText('Qué pidió el cliente, cantidad de láminas, escala, indicaciones del docente, etc.')
    .setRequired(true);

  form.addDateItem().setTitle('Fecha de entrega').setRequired(true);
  form.addTimeItem().setTitle('Hora de entrega');

  form.addMultipleChoiceItem()
    .setTitle('Urgencia')
    .setChoiceValues(['Normal', 'Urgente (menos de 48 h)', 'Express (el mismo día)'])
    .setRequired(true);

  form.addCheckboxItem()
    .setTitle('Formato de entrega')
    .setChoiceValues(['PDF', 'JPG / PNG', 'DWG (AutoCAD)', 'RVT (Revit)', 'SKP (SketchUp)',
                      'Word', 'PowerPoint', 'Excel', 'Físico (impreso o maqueta)'])
    .showOtherOption(true);

  // ---------- SECCIÓN 4: Pago ----------
  form.addPageBreakItem().setTitle('Pago (Yape / Plin / transferencia)');

  var soloNumeros = FormApp.createTextValidation()
    .requireNumberGreaterThanOrEqualTo(0)
    .setHelpText('Solo el número, sin "S/". Ej.: 120')
    .build();

  form.addTextItem().setTitle('Precio total acordado (S/)').setValidation(soloNumeros).setRequired(true);

  form.addMultipleChoiceItem()
    .setTitle('Tipo de pago')
    .setChoiceValues(['Adelanto', 'Pago completo', 'Saldo final', 'Todavía no paga'])
    .setRequired(true);

  form.addTextItem().setTitle('Monto pagado ahora (S/)').setValidation(soloNumeros).setRequired(true);

  form.addMultipleChoiceItem()
    .setTitle('Método de pago')
    .setChoiceValues(['Yape', 'Plin', 'Transferencia bancaria', 'Efectivo'])
    .showOtherOption(true);

  form.addTextItem().setTitle('N.º de operación del Yape / Plin');
  form.addTextItem()
    .setTitle('Nombre que aparece en el Yape')
    .setHelpText('Por si paga otra persona (papá, mamá, amigo).');
  form.addDateItem().setTitle('Fecha del pago');

  // ---------- SECCIÓN 5: Control interno ----------
  form.addPageBreakItem()
    .setTitle('Control interno')
    .setHelpText('Para llevar el seguimiento del pedido.');

  form.addMultipleChoiceItem()
    .setTitle('Estado del pedido')
    .setChoiceValues(['Pendiente', 'En proceso', 'En revisión / correcciones', 'Entregado', 'Cancelado'])
    .setRequired(true);

  form.addTextItem().setTitle('Encargado del trabajo');
  form.addParagraphTextItem().setTitle('Notas / observaciones');

  // ---------- Saltos según el tipo de estudiante ----------
  tipoEstudiante.setChoices([
    tipoEstudiante.createChoice('Estudiante de Arquitectura (1.er a 10.º ciclo)', secArq),
    tipoEstudiante.createChoice('Estudiante de Estudios Generales (hasta 5.º ciclo)', secGen)
  ]);
  // Al terminar la sección de Arquitectura, salta directo al pedido (no pasa por Generales).
  secGen.setGoToPage(secPedido);

  // ---------- Hoja de cálculo que se llena sola ----------
  var hoja = SpreadsheetApp.create(NOMBRE_MARCA + ' - Clientes y pedidos (respuestas)');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, hoja.getId());

  Logger.log('FORMULARIO (para llenar): ' + form.getPublishedUrl());
  Logger.log('FORMULARIO (para editar): ' + form.getEditUrl());
  Logger.log('EXCEL / HOJA DE RESPUESTAS: ' + hoja.getUrl());
}

/** Busca en Drive la imagen "Trampil encabezado" y la pone al inicio. */
function agregarImagenesDeMarca_(form) {
  var archivos = DriveApp.searchFiles(
    "title contains '" + NOMBRE_MARCA + " encabezado' and mimeType contains 'image/' and trashed = false");
  var n = 0;
  while (archivos.hasNext() && n < 1) {
    var archivo = archivos.next();
    form.addImageItem()
      .setImage(archivo.getBlob())
      .setAlignment(FormApp.Alignment.CENTER);
    Logger.log('Imagen agregada: ' + archivo.getName());
    n++;
  }
  if (n === 0) {
    Logger.log('No encontré "' + NOMBRE_MARCA + ' encabezado.png" en tu Drive. ' +
               'Súbelo a Drive y vuelve a ejecutar, o agrégalo a mano.');
  }
}
