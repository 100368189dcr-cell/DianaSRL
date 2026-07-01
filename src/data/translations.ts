export interface TranslationDictionary {
  // Navigation
  agenda: string;
  patients: string;
  odontogram: string;
  payments: string;
  aiAssistant: string;
  syncCenter: string;
  
  // Header / Hero
  clinicName: string;
  clinicTagline: string;
  quickStats: string;
  totalPatients: string;
  pendingAppointments: string;
  completedTreatments: string;
  totalEarnings: string;
  
  // Patients View
  addPatient: string;
  editPatient: string;
  searchPatient: string;
  patientName: string;
  patientPhone: string;
  patientEmail: string;
  patientDni: string;
  patientDob: string;
  patientNotes: string;
  registerPatient: string;
  patientList: string;
  noPatients: string;
  
  // Agenda View
  scheduleAppointment: string;
  selectPatient: string;
  selectDentist: string;
  appointmentDate: string;
  appointmentTime: string;
  treatmentType: string;
  statusScheduled: string;
  statusCompleted: string;
  statusCancelled: string;
  appointmentNotes: string;
  bookNow: string;
  upcomingAppointments: string;
  noAppointments: string;
  
  // Odontograma View
  chartTitle: string;
  chartSubtitle: string;
  selectTooth: string;
  toothState: string;
  stateHealthy: string;
  stateCaries: string;
  stateEndo: string;
  stateMissing: string;
  stateCrown: string;
  stateFilling: string;
  markTreatment: string;
  clinicalRecord: string;
  
  // Payments View
  recordPayment: string;
  amountToPay: string;
  amountPaid: string;
  amountPending: string;
  paymentMethod: string;
  cash: string;
  card: string;
  transfer: string;
  paymentStatus: string;
  invoice: string;
  totalInvoiced: string;
  totalCollected: string;
  pendingReceivable: string;
  billingHistory: string;
  noPayments: string;

  // AI Assistant
  aiTitle: string;
  aiSubtitle: string;
  aiPromptPlaceholder: string;
  aiAnalyze: string;
  aiAnalyzing: string;
  aiResult: string;
  aiResultIntro: string;
  
  // Sync
  syncStatus: string;
  syncActive: string;
  syncDisabled: string;
  googleSheetsUrl: string;
  testSync: string;
  testing: string;
  syncSuccess: string;
  syncError: string;
  syncLogs: string;
  syncNotice: string;
}

export const translations: Record<"en" | "es" | "fr" | "pt", TranslationDictionary> = {
  es: {
    agenda: "Agenda & Citas",
    patients: "Registro de Pacientes",
    odontogram: "Odontograma Clínico",
    payments: "Pagos y Caja",
    aiAssistant: "Asistente Dental IA",
    syncCenter: "Sincronización Sheets",
    
    clinicName: "DianaSRL",
    clinicTagline: "Sistema Odontológico Stetic — Gestión Dental Estética de Alta Fidelidad",
    quickStats: "Estadísticas Generales",
    totalPatients: "Pacientes Registrados",
    pendingAppointments: "Citas Pendientes",
    completedTreatments: "Tratamientos Realizados",
    totalEarnings: "Ingresos Totales",
    
    addPatient: "Agregar Nuevo Paciente",
    editPatient: "Editar Paciente",
    searchPatient: "Buscar por nombre, cédula o teléfono...",
    patientName: "Nombre Completo",
    patientPhone: "Teléfono de Contacto",
    patientEmail: "Correo Electrónico",
    patientDni: "DNI / Cédula de Identidad",
    patientDob: "Fecha de Nacimiento",
    patientNotes: "Notas Médicas / Alergias / Antecedentes",
    registerPatient: "Registrar Paciente en la Base",
    patientList: "Lista de Pacientes",
    noPatients: "No se encontraron pacientes en la base de datos.",
    
    scheduleAppointment: "Agendar Nueva Cita",
    selectPatient: "Seleccionar Paciente",
    selectDentist: "Dentista / Especialista",
    appointmentDate: "Fecha de la Cita",
    appointmentTime: "Hora de la Cita",
    treatmentType: "Tratamiento o Servicio",
    statusScheduled: "Programada",
    statusCompleted: "Completada",
    statusCancelled: "Cancelada",
    appointmentNotes: "Motivo de la consulta / Observaciones",
    bookNow: "Agendar Cita Médica",
    upcomingAppointments: "Próximas Citas Médicas",
    noAppointments: "No hay citas programadas para el periodo seleccionado.",
    
    chartTitle: "Odontograma Clínico Interactivo",
    chartSubtitle: "Haz clic sobre cualquier diente para visualizar su estado actual, registrar diagnósticos y registrar procedimientos estéticos en tiempo real.",
    selectTooth: "Seleccionar Diente",
    toothState: "Estado del Diente",
    stateHealthy: "Sano / Limpio",
    stateCaries: "Caries Activa",
    stateEndo: "Tratamiento de Canal (Endodoncia)",
    stateMissing: "Pieza Ausente (Extracción)",
    stateCrown: "Corona Estética",
    stateFilling: "Resina / Obturación",
    markTreatment: "Aplicar Estado de Pieza",
    clinicalRecord: "Registro Clínico de Pieza",
    
    recordPayment: "Registrar Nuevo Pago",
    amountToPay: "Monto Total del Tratamiento",
    amountPaid: "Monto Pagado",
    amountPending: "Balance Pendiente",
    paymentMethod: "Método de Pago",
    cash: "Efectivo",
    card: "Tarjeta de Crédito / Débito",
    transfer: "Transferencia Bancaria",
    paymentStatus: "Estado del Pago",
    invoice: "Factura",
    totalInvoiced: "Total Facturado",
    totalCollected: "Total Recaudado",
    pendingReceivable: "Por Cobrar Pendiente",
    billingHistory: "Historial de Transacciones de Caja",
    noPayments: "Aún no se han registrado transacciones de pago.",
    
    aiTitle: "Asistente de Diagnóstico Clínico IA",
    aiSubtitle: "Desarrollado con Google Gemini para apoyar tu diagnóstico clínico, sugerir tratamientos estéticos y redactar recomendaciones higiénicas personalizadas.",
    aiPromptPlaceholder: "Escribe los síntomas o hallazgos clínicos del paciente (ej. 'Paciente de 28 años presenta dolor agudo al frío en diente 16, encía inflamada...'): ",
    aiAnalyze: "Generar Recomendación Clínica",
    aiAnalyzing: "Gemini analizando el caso clínico...",
    aiResult: "Sugerencia Clínica Generada",
    aiResultIntro: "Este análisis clínico es un modelo de apoyo asistido por IA para odontólogos de DianaSRL:",
    
    syncStatus: "Estado de la Conexión",
    syncActive: "Sincronización en Vivo Activa",
    syncDisabled: "Sincronización Desactivada",
    googleSheetsUrl: "URL del Web App de Google Apps Script",
    testSync: "Probar Conexión de Sincronización",
    testing: "Enviando ping de prueba a la hoja...",
    syncSuccess: "Sincronización exitosa con la base de Google Sheets.",
    syncError: "Error de conexión. Guardado de forma local de forma segura.",
    syncLogs: "Registro de Actividad de la Base de Datos",
    syncNotice: "Todos los datos registrados en pacientes, citas, odontogramas y cobros se guardan automáticamente en tu navegador (Local Storage) y se sincronizan mediante la API de Google Apps Script con tu hoja de cálculo seleccionada en tiempo real."
  },
  en: {
    agenda: "Calendar & Appointments",
    patients: "Patient Registry",
    odontogram: "Clinical Odontogram",
    payments: "Billing & Cash",
    aiAssistant: "AI Dental Assistant",
    syncCenter: "Google Sheets Sync",
    
    clinicName: "DianaSRL",
    clinicTagline: "Sistema Odontológico Stetic — High-Fidelity Aesthetic Dental Management",
    quickStats: "General Statistics",
    totalPatients: "Registered Patients",
    pendingAppointments: "Pending Bookings",
    completedTreatments: "Completed Procedures",
    totalEarnings: "Total Earnings",
    
    addPatient: "Add New Patient",
    editPatient: "Edit Patient",
    searchPatient: "Search by name, ID, or phone...",
    patientName: "Full Name",
    patientPhone: "Contact Phone",
    patientEmail: "Email Address",
    patientDni: "National ID / Passport",
    patientDob: "Date of Birth",
    patientNotes: "Medical History / Allergies / Background",
    registerPatient: "Register Patient in Database",
    patientList: "Patient Directory",
    noPatients: "No patients found in the database.",
    
    scheduleAppointment: "Schedule New Appointment",
    selectPatient: "Select Patient",
    selectDentist: "Dentist / Specialist",
    appointmentDate: "Appointment Date",
    appointmentTime: "Appointment Time",
    treatmentType: "Treatment or Service",
    statusScheduled: "Scheduled",
    statusCompleted: "Completed",
    statusCancelled: "Cancelled",
    appointmentNotes: "Reason for visit / Clinical notes",
    bookNow: "Book Medical Appointment",
    upcomingAppointments: "Upcoming Dental Appointments",
    noAppointments: "No appointments scheduled for the selected period.",
    
    chartTitle: "Interactive Clinical Odontogram",
    chartSubtitle: "Click on any tooth to view its current state, register clinical findings, and mark aesthetic procedures in real time.",
    selectTooth: "Select Tooth",
    toothState: "Tooth Condition",
    stateHealthy: "Healthy / Clean",
    stateCaries: "Active Caries",
    stateEndo: "Root Canal (Endodontics)",
    stateMissing: "Missing Tooth (Extracted)",
    stateCrown: "Aesthetic Crown",
    stateFilling: "Composite / Filling",
    markTreatment: "Apply Tooth Status",
    clinicalRecord: "Clinical Tooth History",
    
    recordPayment: "Record New Payment",
    amountToPay: "Total Treatment Cost",
    amountPaid: "Amount Paid",
    amountPending: "Remaining Balance",
    paymentMethod: "Payment Method",
    cash: "Cash",
    card: "Credit / Debit Card",
    transfer: "Bank Transfer",
    paymentStatus: "Payment Status",
    invoice: "Invoice",
    totalInvoiced: "Total Invoiced",
    totalCollected: "Total Collected",
    pendingReceivable: "Accounts Receivable",
    billingHistory: "Cash Transaction Ledger",
    noPayments: "No payment transactions have been recorded yet.",
    
    aiTitle: "AI Clinical Diagnostics Assistant",
    aiSubtitle: "Powered by Google Gemini to assist with clinical diagnosis, suggest aesthetic treatments, and draft personalized oral care advice.",
    aiPromptPlaceholder: "Type patient symptoms or clinical findings (e.g., '28-year-old patient reports acute cold sensitivity on tooth 16, swollen gums...'): ",
    aiAnalyze: "Generate Clinical Suggestion",
    aiAnalyzing: "Gemini analyzing the clinical case...",
    aiResult: "Generated Clinical Diagnosis Support",
    aiResultIntro: "This clinical evaluation is an AI-powered assistant model for DianaSRL dentists:",
    
    syncStatus: "Connection Status",
    syncActive: "Live Sync Active",
    syncDisabled: "Sync Disabled",
    googleSheetsUrl: "Google Apps Script Web App URL",
    testSync: "Test Synchronization Link",
    testing: "Sending ping to Google Spreadsheet...",
    syncSuccess: "Successfully synchronized with Google Sheets database.",
    syncError: "Connection error. Stored securely in local cache.",
    syncLogs: "Database Activity Logs",
    syncNotice: "All registered data for patients, appointments, dental charts, and billing is saved automatically in your browser (Local Storage) and synced to your Google Spreadsheet via the Apps Script Web App API."
  },
  fr: {
    agenda: "Agenda & Rendez-vous",
    patients: "Registre des Patients",
    odontogram: "Odontogramme Clinique",
    payments: "Paiements & Caisse",
    aiAssistant: "Assistant Dentaire IA",
    syncCenter: "Synchro Google Sheets",
    
    clinicName: "DianaSRL",
    clinicTagline: "Sistema Odontológico Stetic — Gestion Dentaire Esthétique Haute Fidélité",
    quickStats: "Statistiques Générales",
    totalPatients: "Patients Enregistrés",
    pendingAppointments: "Rendez-vous en Attente",
    completedTreatments: "Traitements Effectués",
    totalEarnings: "Revenus Totaux",
    
    addPatient: "Ajouter un Patient",
    editPatient: "Modifier le Patient",
    searchPatient: "Rechercher par nom, DNI ou téléphone...",
    patientName: "Nom Complet",
    patientPhone: "Téléphone de Contact",
    patientEmail: "Adresse Email",
    patientDni: "Numéro de Carte d'Identité / Passeport",
    patientDob: "Date de Naissance",
    patientNotes: "Notes Médicales / Allergies / Antécédents",
    registerPatient: "Enregistrer le Patient",
    patientList: "Répertoire des Patients",
    noPatients: "Aucun patient trouvé dans la base de données.",
    
    scheduleAppointment: "Prendre Rendez-vous",
    selectPatient: "Sélectionner le Patient",
    selectDentist: "Dentiste / Spécialiste",
    appointmentDate: "Date du Rendez-vous",
    appointmentTime: "Heure du Rendez-vous",
    treatmentType: "Traitement ou Service",
    statusScheduled: "Planifié",
    statusCompleted: "Complété",
    statusCancelled: "Annulé",
    appointmentNotes: "Motif de consultation / Observations",
    bookNow: "Prendre Rendez-vous",
    upcomingAppointments: "Rendez-vous Dentaires À Venir",
    noAppointments: "Aucun rendez-vous prévu pour la période sélectionnée.",
    
    chartTitle: "Odontogramme Clinique Interactif",
    chartSubtitle: "Cliquez sur une dent pour afficher son état, inscrire un diagnostic ou enregistrer un traitement esthétique en temps réel.",
    selectTooth: "Sélectionner la Dent",
    toothState: "État de la Dent",
    stateHealthy: "Saine / Propre",
    stateCaries: "Carie Active",
    stateEndo: "Traitement de Canal (Endodontie)",
    stateMissing: "Dent Absente (Extraction)",
    stateCrown: "Couronne Esthétique",
    stateFilling: "Résine / Obturation",
    markTreatment: "Appliquer l'État",
    clinicalRecord: "Dossier Clinique de la Dent",
    
    recordPayment: "Enregistrer un Paiement",
    amountToPay: "Montant Total du Traitement",
    amountPaid: "Montant Payé",
    amountPending: "Solde Restant",
    paymentMethod: "Mode de Paiement",
    cash: "Espèces",
    card: "Carte de Crédit / Débit",
    transfer: "Virement Bancaire",
    paymentStatus: "Statut du Paiement",
    invoice: "Facture",
    totalInvoiced: "Total Facturé",
    totalCollected: "Total Collecté",
    pendingReceivable: "Créances Client",
    billingHistory: "Historique des Transactions de Caisse",
    noPayments: "Aucune transaction enregistrée pour le moment.",
    
    aiTitle: "Assistant IA de Diagnostic Clinique",
    aiSubtitle: "Propulsé par Google Gemini pour vous assister dans vos diagnostics, suggérer des soins esthétiques et rédiger des conseils d'hygiène personnalisés.",
    aiPromptPlaceholder: "Saisissez les symptômes ou les observations cliniques (ex: 'Patient de 28 ans présentant une sensibilité aiguë au froid sur la dent 16, gencive gonflée...'): ",
    aiAnalyze: "Générer la Recommandation Clinique",
    aiAnalyzing: "Gemini analyse le cas clinique...",
    aiResult: "Diagnostic IA Généré",
    aiResultIntro: "Cette évaluation clinique est un modèle d'aide assisté par l'IA pour les dentistes de DianaSRL:",
    
    syncStatus: "Statut de la Connexion",
    syncActive: "Synchro en Temps Réel Active",
    syncDisabled: "Synchro Désactivée",
    googleSheetsUrl: "URL Web App Google Apps Script",
    testSync: "Tester le Lien de Sincronisation",
    testing: "Ping en cours vers Google Sheets...",
    syncSuccess: "Sincronisation réussie avec la feuille de calcul Google Sheets.",
    syncError: "Erreur de connexion. Sauvegardé en local de manière sécurisée.",
    syncLogs: "Historique des Activités de la Base de Datos",
    syncNotice: "Toutes vos données (patients, rendez-vous, odontogrammes et factures) sont stockées dans le cache local de votre navigateur et synchronisées en temps réel vers votre Google Sheet via l'API Web App."
  },
  pt: {
    agenda: "Agenda & Consultas",
    patients: "Cadastro de Pacientes",
    odontogram: "Odontograma Clínico",
    payments: "Pagos e Caixa",
    aiAssistant: "Assistente Odontológico IA",
    syncCenter: "Sincronização Sheets",
    
    clinicName: "DianaSRL",
    clinicTagline: "Sistema Odontológico Stetic — Gestão Odontológica Estética de Alta Fidelidade",
    quickStats: "Estatísticas Gerais",
    totalPatients: "Pacientes Registrados",
    pendingAppointments: "Consultas Pendentes",
    completedTreatments: "Tratamentos Realizados",
    totalEarnings: "Receita Total",
    
    addPatient: "Cadastrar Novo Paciente",
    editPatient: "Editar Paciente",
    searchPatient: "Buscar por nome, CPF/DNI ou telefone...",
    patientName: "Nome Completo",
    patientPhone: "Telefone de Contato",
    patientEmail: "E-mail",
    patientDni: "DNI / CPF de Identidade",
    patientDob: "Data de Nascimento",
    patientNotes: "Histórico Médico / Alergias / Observações",
    registerPatient: "Salvar Paciente no Banco",
    patientList: "Lista de Pacientes",
    noPatients: "Nenhum paciente cadastrado no banco de dados.",
    
    scheduleAppointment: "Agendar Nova Consulta",
    selectPatient: "Selecionar Paciente",
    selectDentist: "Dentista / Especialista",
    appointmentDate: "Data da Consulta",
    appointmentTime: "Horário da Consulta",
    treatmentType: "Tratamento ou Serviço",
    statusScheduled: "Programada",
    statusCompleted: "Concluída",
    statusCancelled: "Cancelada",
    appointmentNotes: "Motivo da consulta / Observações clínicas",
    bookNow: "Agendar Consulta",
    upcomingAppointments: "Próximas Consultas Médicas",
    noAppointments: "Nenhuma consulta agendada para o período selecionado.",
    
    chartTitle: "Odontograma Clínico Interativo",
    chartSubtitle: "Clique em qualquer dente para visualizar seu status, registrar diagnósticos e acompanhar tratamentos estéticos em tempo real.",
    selectTooth: "Selecionar Dente",
    toothState: "Estado do Dente",
    stateHealthy: "Saudável / Limpo",
    stateCaries: "Cárie Ativa",
    stateEndo: "Tratamento de Canal (Endodontia)",
    stateMissing: "Elemento Ausente (Extraído)",
    stateCrown: "Coroa Estética",
    stateFilling: "Resina / Obturação",
    markTreatment: "Aplicar Estado da Peça",
    clinicalRecord: "Registro Clínico do Dente",
    
    recordPayment: "Registrar Novo Pagamento",
    amountToPay: "Valor Total do Tratamiento",
    amountPaid: "Valor Pago",
    amountPending: "Saldo Devedor",
    paymentMethod: "Forma de Pagamento",
    cash: "Dinheiro / Pix",
    card: "Cartão de Crédito / Débito",
    transfer: "Transferência Bancária",
    paymentStatus: "Status do Pagamento",
    invoice: "Recibo / Fatura",
    totalInvoiced: "Total Faturado",
    totalCollected: "Total Recebido",
    pendingReceivable: "Contas a Receber",
    billingHistory: "Histórico de Transações de Caixa",
    noPayments: "Nenhum pagamento registrado até o momento.",
    
    aiTitle: "Assistente IA de Diagnóstico Clínico",
    aiSubtitle: "Desenvolvido com o Google Gemini para auxiliar na análise clínica, sugerir tratamentos estéticos modernos e criar rotinas de higiene oral personalizadas.",
    aiPromptPlaceholder: "Digite as queixas ou achados do paciente (ex: 'Paciente de 28 anos relata sensibilidade ao frio no dente 16, gengiva sangrando...'): ",
    aiAnalyze: "Gerar Parecer Clínico IA",
    aiAnalyzing: "Gemini analisando o caso clínico...",
    aiResult: "Sugestão Clínica Gerada",
    aiResultIntro: "Este diagnóstico é um suporte assistido por Inteligência Artificial para os odontologistas da DianaSRL:",
    
    syncStatus: "Status da Conexão",
    syncActive: "Sincronização em Tempo Real Ativa",
    syncDisabled: "Sincronização Desativada",
    googleSheetsUrl: "URL do Web App do Google Apps Script",
    testSync: "Testar Conexão de Sincronização",
    testing: "Enviando ping para o Google Sheets...",
    syncSuccess: "Sincronização realizada com sucesso no Google Sheets.",
    syncError: "Erro de conexão. Gravado localmente em segurança.",
    syncLogs: "Registro de Atividade da Base de Dados",
    syncNotice: "Todos os dados de pacientes, consultas, prontuários de dentes e faturamento são salvos de forma local no navegador (Local Storage) e transmitidos via API do Google Apps Script para sua planilha de forma automática."
  }
};
