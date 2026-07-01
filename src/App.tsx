import React, { useState, useEffect } from "react";
import { 
  Calendar, 
  Users, 
  Activity, 
  DollarSign, 
  Sparkles, 
  Database, 
  Moon, 
  Sun, 
  Globe, 
  Plus, 
  Search, 
  CheckCircle, 
  XCircle, 
  PlusCircle, 
  Edit, 
  Trash2, 
  Heart, 
  TrendingUp, 
  Clock, 
  Printer, 
  ArrowRight, 
  LogOut, 
  CreditCard, 
  AlertTriangle, 
  FileText, 
  Layers,
  ChevronRight,
  Info,
  Check,
  RefreshCw,
  Lock,
  Shield,
  Download,
  Palette,
  Settings,
  ToggleLeft,
  ToggleRight,
  ShieldCheck,
  Eye,
  EyeOff,
  Copy
} from "lucide-react";
import { motion, AnimatePresence } from "motion/react";
import { translations } from "./data/translations";
import { Patient, Appointment, ToothCondition, PaymentRecord, SyncLog, AdminUser, DoctorAvailability, PaymentMethodConfig } from "./types";
// @ts-ignore
import clinicLogo from "./assets/images/clinic_favicon_1782900343266.jpg";

const APPS_SCRIPT_CODE = `function doGet(e) {
  var spreadsheetUrl = e.parameter.spreadsheetUrl || "https://docs.google.com/spreadsheets/d/1KS9ngWCTTZPfT0Tr8rHBLWz2YVFFH57AHGtx9J_iZio/edit";
  return handleRequest({ action: "read_all", spreadsheetUrl: spreadsheetUrl });
}

function doPost(e) {
  var request = JSON.parse(e.postData.contents);
  return handleRequest(request);
}

function handleRequest(request) {
  var spreadsheetUrl = request.spreadsheetUrl || "https://docs.google.com/spreadsheets/d/1KS9ngWCTTZPfT0Tr8rHBLWz2YVFFH57AHGtx9J_iZio/edit";
  var ss;
  try {
    ss = SpreadsheetApp.openByUrl(spreadsheetUrl);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ 
      success: false, 
      error: "Error al abrir la hoja de cálculo por URL: " + err.message + ". Asegúrate de compartir la hoja de cálculo con permisos de edición." 
    })).setMimeType(ContentService.MimeType.JSON);
  }
  var action = request.action;
  
  if (action === "read_all" || action === "ping") {
    if (action === "ping") {
      return ContentService.createTextOutput(JSON.stringify({ success: true, ping: "pong" })).setMimeType(ContentService.MimeType.JSON);
    }
    var data = {
      patients: getSheetData(ss, "Pacientes"),
      appointments: getSheetData(ss, "Citas"),
      payments: getSheetData(ss, "Pagos"),
      odontograms: getSheetData(ss, "Odontogramas")
    };
    return ContentService.createTextOutput(JSON.stringify({ success: true, data: data }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  if (action === "bulk_write") {
    var payloadData = request.data || {};
    var sheetNames = Object.keys(payloadData);
    
    for (var s = 0; s < sheetNames.length; s++) {
      var sName = sheetNames[s];
      var rowsData = payloadData[sName];
      if (!Array.isArray(rowsData)) continue;
      
      var sheet = ss.getSheetByName(sName);
      if (!sheet) {
        sheet = ss.insertSheet(sName);
      }
      sheet.clear();
      if (rowsData.length === 0) continue;
      
      var headers = Object.keys(rowsData[0]);
      sheet.appendRow(headers);
      
      for (var r = 0; r < rowsData.length; r++) {
        var rowObj = rowsData[r];
        var row = [];
        for (var h = 0; h < headers.length; h++) {
          var val = rowObj[headers[h]];
          row.push(val !== undefined && val !== null ? String(val) : "");
        }
        sheet.appendRow(row);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({ success: true, message: "Carga masiva completada con éxito." }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  var sheet = ss.getSheetByName(action || "Pacientes");
  if (!sheet) {
    sheet = ss.insertSheet(action || "Pacientes");
  }
  
  var payloadData = request.data || {};
  var headers = sheet.getLastRow() === 0 ? [] : sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var keys = Object.keys(payloadData);
  
  if (headers.length === 0) {
    headers = keys;
    sheet.appendRow(headers);
  } else {
    for (var i = 0; i < keys.length; i++) {
      if (headers.indexOf(keys[i]) === -1) {
        headers.push(keys[i]);
        sheet.getRange(1, headers.length).setValue(keys[i]);
      }
    }
  }
  
  var row = [];
  for (var i = 0; i < headers.length; i++) {
    row.push(payloadData[headers[i]] !== undefined ? payloadData[headers[i]] : "");
  }
  sheet.appendRow(row);
  
  return ContentService.createTextOutput(JSON.stringify({ success: true, row: row }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheetData(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) return [];
  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  var headers = values[0];
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var row = {};
    for (var j = 0; j < headers.length; j++) {
      var val = values[i][j];
      row[headers[j]] = val;
    }
    rows.push(row);
  }
  return rows;
}`;

export default function App() {
  // Navigation State
  const [activeTab, setActiveTab] = useState<"agenda" | "patients" | "odontogram" | "payments" | "ai" | "sync" | "availability" | "master-config">("agenda");
  const [copiedScript, setCopiedScript] = useState(false);
  
  // Portal Mode State ("admin" for Dentist side, "patient" for Client side)
  const [portalMode, setPortalMode] = useState<"admin" | "patient">(() => {
    const params = new URLSearchParams(window.location.search);
    const portal = params.get("portal") || params.get("mode");
    if (portal === "admin") return "admin";
    if (portal === "patient" || portal === "client") return "patient";
    return "patient"; // Por defecto, es el portal del paciente
  });

  const changePortalMode = (mode: "admin" | "patient") => {
    setPortalMode(mode);
    const newUrl = window.location.pathname + `?portal=${mode}`;
    window.history.pushState(null, "", newUrl);
    triggerToast(
      lang === "es" 
        ? (mode === "admin" ? "Accediendo al Portal Clínico Administrador" : "Accediendo al Portal de Pacientes") 
        : (mode === "admin" ? "Accessing Clinical Admin Portal" : "Accessing Patient Portal"), 
      "success"
    );
  };

  const [loggedPatientId, setLoggedPatientId] = useState<string>(() => {
    return localStorage.getItem("dianasrl_logged_patient_id") || "";
  });
  const [patientActiveTab, setPatientActiveTab] = useState<"my-appointments" | "my-odontogram" | "my-billing" | "my-ai">("my-appointments");

  // Client Identification State
  const [patientDniInput, setPatientDniInput] = useState<string>("");
  const [patientRegForm, setPatientRegForm] = useState({
    name: "",
    phone: "",
    email: "",
    dni: "",
    dob: "",
    notes: "",
    allergies: "",
    medicalConditions: "",
    consultationReason: "",
    dentalInsurance: "",
    preferredContact: "WhatsApp",
    address: ""
  });

  // Client Appointment Request State
  const [patientNewAppt, setPatientNewAppt] = useState({
    dentistName: "Dra. Diana Rojas",
    date: "",
    time: "",
    treatmentType: "",
    notes: ""
  });

  // Client AI Companion State
  const [patientAiPrompt, setPatientAiPrompt] = useState<string>("");
  const [patientAiResponse, setPatientAiResponse] = useState<string>("");
  const [isPatientAiLoading, setIsPatientAiLoading] = useState<boolean>(false);

  // Client view-only selected tooth state
  const [patientSelectedTooth, setPatientSelectedTooth] = useState<number | null>(null);

  // Admin Database & Session States
  const [admins, setAdmins] = useState<AdminUser[]>([]);
  const [currentAdmin, setCurrentAdmin] = useState<AdminUser | null>(() => {
    const saved = localStorage.getItem("dianasrl_current_admin");
    return saved ? JSON.parse(saved) : null;
  });
  const [adminUsernameInput, setAdminUsernameInput] = useState<string>("");
  const [adminPasswordInput, setAdminPasswordInput] = useState<string>("");

  // New admin registration state (exclusive for Master founder account)
  const [newAdminForm, setNewAdminForm] = useState({
    name: "",
    username: "",
    password: ""
  });

  // Admin Login Action
  const handleAdminLogin = (e?: React.FormEvent) => {
    if (e) e.preventDefault();
    const usernameClean = adminUsernameInput.trim().toLowerCase();
    const passwordClean = adminPasswordInput.trim();

    if (!usernameClean || !passwordClean) {
      triggerToast("Por favor ingrese su usuario y contraseña.", "error");
      return;
    }

    const foundAdmin = admins.find(
      a => a.username.toLowerCase() === usernameClean && a.passwordHash === passwordClean
    );

    if (foundAdmin) {
      setCurrentAdmin(foundAdmin);
      localStorage.setItem("dianasrl_current_admin", JSON.stringify(foundAdmin));
      triggerToast(`¡Acceso concedido! Bienvenido, ${foundAdmin.name}.`, "success");
      setAdminUsernameInput("");
      setAdminPasswordInput("");
    } else {
      triggerToast("Usuario o contraseña incorrectos. Verifique sus credenciales.", "error");
    }
  };

  // Admin Logout Action
  const handleAdminLogout = () => {
    setCurrentAdmin(null);
    localStorage.removeItem("dianasrl_current_admin");
    triggerToast("Sesión de administrador cerrada correctamente.", "success");
  };

  // Admin Register Action (Only accessible by Master admin role)
  const handleRegisterAdmin = (e: React.FormEvent) => {
    e.preventDefault();
    
    // Safety check
    if (!currentAdmin || currentAdmin.role !== "master") {
      triggerToast("Acceso denegado. Solo la cuenta fundadora (Master) puede registrar nuevos administradores.", "error");
      return;
    }

    const nameTrim = newAdminForm.name.trim();
    const userTrim = newAdminForm.username.trim().toLowerCase();
    const passTrim = newAdminForm.password.trim();

    if (!nameTrim || !userTrim || !passTrim) {
      triggerToast("Todos los campos de registro son obligatorios.", "error");
      return;
    }

    // Check if user already exists
    const exists = admins.some(a => a.username.toLowerCase() === userTrim);
    if (exists) {
      triggerToast(`El usuario "${userTrim}" ya se encuentra registrado.`, "error");
      return;
    }

    const newAdmin: AdminUser = {
      id: "adm_" + Math.random().toString(36).substring(2, 9),
      name: nameTrim,
      username: userTrim,
      passwordHash: passTrim, // Simple local-first plain password storage representation
      role: "admin",
      createdAt: new Date().toISOString()
    };

    const updatedAdmins = [...admins, newAdmin];
    setAdmins(updatedAdmins);
    localStorage.setItem("dianasrl_admins", JSON.stringify(updatedAdmins));

    // Reset registration form
    setNewAdminForm({ name: "", username: "", password: "" });
    triggerToast(`Administrador "${nameTrim}" registrado con éxito.`, "success");
  };


  // Patient Login Action
  const handlePatientLogin = (dniOrName: string) => {
    const cleanTerm = dniOrName.trim().toLowerCase();
    if (!cleanTerm) {
      triggerToast("Por favor ingrese su cédula, DNI o nombre para buscar.", "error");
      return;
    }
    const found = patients.find(p => 
      p.dni.replace(/[^a-zA-Z0-9]/g, "").includes(cleanTerm.replace(/[^a-zA-Z0-9]/g, "")) || 
      p.name.toLowerCase().includes(cleanTerm)
    );
    if (found) {
      setLoggedPatientId(found.id);
      localStorage.setItem("dianasrl_logged_patient_id", found.id);
      triggerToast(lang === "es" ? `¡Bienvenido, ${found.name}!` : `Welcome, ${found.name}!`, "success");
      setPatientDniInput("");
    } else {
      triggerToast(lang === "es" ? "No encontramos ningún registro de paciente con esos datos." : "No patient record found with these details.", "error");
    }
  };

  // Patient Register Action
  const handlePatientRegister = (e: React.FormEvent) => {
    e.preventDefault();
    if (!patientRegForm.name || !patientRegForm.dni) {
      triggerToast(lang === "es" ? "El nombre y la identificación (Cédula/DNI/Pasaporte) son obligatorios." : "Name and Identification (DNI/ID/Passport) are required.", "error");
      return;
    }

    const exists = patients.find(p => p.dni.replace(/[^a-zA-Z0-9]/g, "").toLowerCase() === patientRegForm.dni.replace(/[^a-zA-Z0-9]/g, "").toLowerCase());
    if (exists) {
      triggerToast(lang === "es" ? "Ya existe un paciente registrado con este documento (Cédula/DNI/Pasaporte)." : "A patient is already registered with this document/ID/Passport.", "error");
      return;
    }

    const created: Patient = {
      id: "p_" + Math.random().toString(36).substring(2, 9),
      name: patientRegForm.name,
      phone: patientRegForm.phone || "",
      email: patientRegForm.email || "",
      dni: patientRegForm.dni,
      dob: patientRegForm.dob || "",
      notes: patientRegForm.notes || "",
      allergies: patientRegForm.allergies || "",
      medicalConditions: patientRegForm.medicalConditions || "",
      consultationReason: patientRegForm.consultationReason || "",
      dentalInsurance: patientRegForm.dentalInsurance || "",
      preferredContact: patientRegForm.preferredContact || "WhatsApp",
      address: patientRegForm.address || "",
      registeredAt: new Date().toISOString()
    };

    const updatedList = [created, ...patients];
    setPatients(updatedList);
    localStorage.setItem("dianasrl_patients", JSON.stringify(updatedList));
    setLoggedPatientId(created.id);
    localStorage.setItem("dianasrl_logged_patient_id", created.id);

    // Sync to Sheets
    executeSync("Pacientes", created);

    // Reset Form
    setPatientRegForm({
      name: "",
      phone: "",
      email: "",
      dni: "",
      dob: "",
      notes: "",
      allergies: "",
      medicalConditions: "",
      consultationReason: "",
      dentalInsurance: "",
      preferredContact: "WhatsApp",
      address: ""
    });
    triggerToast(lang === "es" ? "¡Registro exitoso! Bienvenido a DianaSRL." : "Registration successful! Welcome to DianaSRL.", "success");
  };

  // Patient Book Appointment Action
  const handlePatientRequestAppt = (e: React.FormEvent) => {
    e.preventDefault();
    if (!loggedPatientId) return;
    const patient = patients.find(p => p.id === loggedPatientId);
    if (!patient) return;

    if (!patientNewAppt.date || !patientNewAppt.time || !patientNewAppt.treatmentType) {
      triggerToast("Por favor complete todos los campos requeridos (*).", "error");
      return;
    }

    const created: Appointment = {
      id: "a_" + Math.random().toString(36).substring(2, 9),
      patientId: patient.id,
      patientName: patient.name,
      dentistName: patientNewAppt.dentistName || "Dra. Diana Rojas",
      date: patientNewAppt.date,
      time: patientNewAppt.time,
      treatmentType: patientNewAppt.treatmentType,
      notes: patientNewAppt.notes || "",
      status: "Scheduled"
    };

    const updatedList = [created, ...appointments];
    setAppointments(updatedList);
    localStorage.setItem("dianasrl_appointments", JSON.stringify(updatedList));

    // Trigger Sheet sync
    executeSync("Citas", created);

    // Reset Form
    setPatientNewAppt({
      dentistName: "Dra. Diana Rojas",
      date: "",
      time: "",
      treatmentType: "",
      notes: ""
    });
    triggerToast(lang === "es" ? "¡Su cita médica ha sido agendada correctamente!" : "Your appointment has been scheduled successfully!", "success");
  };

  // Save Doctor Availability Action
  const handleSaveAvailability = (e: React.FormEvent) => {
    e.preventDefault();
    if (!availDate) {
      triggerToast(lang === "es" ? "Por favor seleccione una fecha." : "Please select a date.", "error");
      return;
    }
    const docName = availDoctor || (currentAdmin ? currentAdmin.name : "Dra. Diana Rojas");
    if (!docName) {
      triggerToast(lang === "es" ? "Por favor seleccione un doctor." : "Please select a doctor.", "error");
      return;
    }
    if (availSlots.length === 0) {
      triggerToast(lang === "es" ? "Debe seleccionar al menos una hora disponible." : "Please select at least one available hour.", "error");
      return;
    }

    // Check if availability already exists for this doctor and date
    const existingIndex = availabilities.findIndex(
      a => a.doctorName.toLowerCase() === docName.toLowerCase() && a.date === availDate
    );

    let updatedList = [...availabilities];

    if (existingIndex !== -1) {
      // Update
      updatedList[existingIndex] = {
        ...updatedList[existingIndex],
        slots: [...availSlots].sort()
      };
    } else {
      // Insert new
      const newAvail: DoctorAvailability = {
        id: "av_" + Math.random().toString(36).substring(2, 9),
        doctorName: docName,
        date: availDate,
        slots: [...availSlots].sort()
      };
      updatedList.push(newAvail);
    }

    setAvailabilities(updatedList);
    localStorage.setItem("dianasrl_availabilities", JSON.stringify(updatedList));

    // Reset slots/date
    setAvailSlots([]);
    setAvailDate("");

    triggerToast(
      lang === "es" 
        ? `Disponibilidad guardada correctamente para ${docName}.` 
        : `Availability successfully saved for ${docName}.`, 
      "success"
    );
  };

  // Delete Doctor Availability Action
  const handleDeleteAvailability = (id: string) => {
    const updatedList = availabilities.filter(a => a.id !== id);
    setAvailabilities(updatedList);
    localStorage.setItem("dianasrl_availabilities", JSON.stringify(updatedList));
    triggerToast(
      lang === "es" ? "Disponibilidad de fecha eliminada." : "Date availability deleted.", 
      "success"
    );
  };

  // Patient AI Advice Action
  const handlePatientAiAnalyze = async () => {
    if (!patientAiPrompt.trim()) return;
    setIsPatientAiLoading(true);
    setPatientAiResponse("");

    try {
      const response = await fetch("/api/dental-advisor", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ symptoms: patientAiPrompt, lang: lang })
      });
      const data = await response.json();
      if (data.success) {
        setPatientAiResponse(data.analysis);
      } else {
        setPatientAiResponse("Fallo al contactar el servicio de Inteligencia Artificial.");
      }
    } catch (e) {
      setPatientAiResponse("Ocurrió un error al contactar al servidor clínico.");
    } finally {
      setIsPatientAiLoading(false);
    }
  };

  // Patient Logout
  const handlePatientLogout = () => {
    setLoggedPatientId("");
    localStorage.removeItem("dianasrl_logged_patient_id");
    triggerToast(lang === "es" ? "Sesión del portal de pacientes cerrada." : "Patient portal session closed.", "success");
  };
  
  // Theme and Language
  const [lang, setLang] = useState<"en" | "es" | "fr" | "pt">("es");
  const [theme, setTheme] = useState<"light" | "dark">("light");

  const [scriptUrl, setScriptUrl] = useState<string>(() => {
    const saved = localStorage.getItem("dianasrl_script_url");
    if (!saved || saved.includes("AKfycbwZYD4fiYPh")) {
      return "https://script.google.com/macros/s/AKfycbyDnTDgcn54bJ9koOhwHMfUeVpq0uj9FmBkdExneZy6jUSCdgV3vNDzTyoMwwHSWz3yyg/exec";
    }
    return saved;
  });
  const [spreadsheetUrl, setSpreadsheetUrl] = useState<string>(() => {
    return localStorage.getItem("dianasrl_spreadsheet_url") || "https://docs.google.com/spreadsheets/d/1KS9ngWCTTZPfT0Tr8rHBLWz2YVFFH57AHGtx9J_iZio/edit?gid=1256709280#gid=1256709280";
  });
  const [liveSync, setLiveSync] = useState<boolean>(() => {
    const saved = localStorage.getItem("dianasrl_live_sync");
    return saved !== null ? JSON.parse(saved) : true;
  });

  // Core App Databases
  const [patients, setPatients] = useState<Patient[]>([]);
  const [appointments, setAppointments] = useState<Appointment[]>([]);
  const [selectedPatientId, setSelectedPatientId] = useState<string>("");
  const [toothConditions, setToothConditions] = useState<Record<string, ToothCondition[]>>({});
  const [payments, setPayments] = useState<PaymentRecord[]>([]);
  const [gateways, setGateways] = useState<PaymentMethodConfig[]>(() => {
    const saved = localStorage.getItem("dianasrl_gateways");
    if (saved) {
      try {
        return JSON.parse(saved);
      } catch (e) {}
    }
    return [
      {
        id: "banco_popular",
        name: "Banco Popular Dominicano",
        enabled: true,
        instructions: "Realice transferencia o depósito al número de cuenta corriente 796541234 a nombre de DianaSRL, luego envíe su comprobante por WhatsApp.",
        link: "https://www.popularenlinea.com"
      },
      {
        id: "banco_bhd",
        name: "Banco BHD León",
        enabled: true,
        instructions: "Realice transferencia al número de cuenta de ahorros 892345678 a nombre de DianaSRL y envíe comprobante de pago.",
        link: "https://www.bhd.com.do"
      },
      {
        id: "paypal",
        name: "PayPal Internacional",
        enabled: true,
        instructions: "Pague de forma rápida usando su balance PayPal o tarjetas internacionales en el link directo.",
        link: "https://paypal.me/dianasrl"
      },
      {
        id: "stripe_card",
        name: "Pago con Tarjeta de Crédito (Stripe Link)",
        enabled: false,
        instructions: "Ingrese al enlace para realizar su pago inmediato con tarjeta Visa/Mastercard.",
        link: "https://buy.stripe.com/mock-dianasrl"
      }
    ];
  });
  const [syncLogs, setSyncLogs] = useState<SyncLog[]>([]);
  const [availabilities, setAvailabilities] = useState<DoctorAvailability[]>([]);
  const [paymentsSubTab, setPaymentsSubTab] = useState<"caja" | "config">("caja");

  // Selection states for sub-views
  const [selectedTooth, setSelectedTooth] = useState<number | null>(null);
  const [selectedInvoice, setSelectedInvoice] = useState<PaymentRecord | null>(null);

  // Form submission states
  const [newPatient, setNewPatient] = useState<Partial<Patient>>({ name: "", phone: "", email: "", dni: "", dob: "", notes: "" });
  const [newAppointment, setNewAppointment] = useState<Partial<Appointment>>({ patientId: "", dentistName: "", date: "", time: "", treatmentType: "", notes: "" });
  const [newPayment, setNewPayment] = useState<Partial<PaymentRecord>>({ patientId: "", treatmentType: "", amountTotal: 0, amountPaid: 0, paymentMethod: "Cash" });

  // Availability management states
  const [availDate, setAvailDate] = useState<string>("");
  const [availDoctor, setAvailDoctor] = useState<string>("");
  const [availSlots, setAvailSlots] = useState<string[]>([]);

  // AI Assistant states
  const [aiPrompt, setAiPrompt] = useState<string>("");
  const [aiResponse, setAiResponse] = useState<string>("");
  const [isAiLoading, setIsAiLoading] = useState<boolean>(false);

  // Feedback notifications
  const [toastMessage, setToastMessage] = useState<{ type: "success" | "error"; text: string } | null>(null);
  const [isTestingSync, setIsTestingSync] = useState<boolean>(false);
  const [isImportingSync, setIsImportingSync] = useState<boolean>(false);
  const [isExportingSync, setIsExportingSync] = useState<boolean>(false);
  const [showLangDropdown, setShowLangDropdown] = useState<boolean>(false);
  const [showAdminNavDropdown, setShowAdminNavDropdown] = useState<boolean>(false);

  // Master Dynamic Configuration States
  const [clinicName, setClinicName] = useState<string>(() => {
    return localStorage.getItem("dianasrl_clinic_name") || "DianaSRL";
  });
  const [clinicSlogan, setClinicSlogan] = useState<string>(() => {
    return localStorage.getItem("dianasrl_clinic_slogan") || "Odontología Estética de Vanguardia";
  });
  const [clinicColorTheme, setClinicColorTheme] = useState<"pink" | "emerald" | "amber" | "violet" | "sky">(() => {
    return (localStorage.getItem("dianasrl_clinic_color_theme") as any) || "pink";
  });
  const [activeFeatures, setActiveFeatures] = useState<Record<string, boolean>>(() => {
    const saved = localStorage.getItem("dianasrl_active_features");
    if (saved) {
      try { return JSON.parse(saved); } catch (e) {}
    }
    return {
      aiDiagnostic: true,
      payments: true,
      liveSync: true,
      onlineRequests: true,
    };
  });
  const [termsAndConditions, setTermsAndConditions] = useState<string>(() => {
    return localStorage.getItem("dianasrl_terms_conditions") || 
      `TÉRMINOS Y CONDICIONES GENERALES\n\n1. Servicios Ofrecidos: Esta clínica ofrece tratamientos de odontología estética, endodoncia, ortodoncia e implantes. Los análisis complementarios de IA son únicamente de pre-diagnóstico y orientación.\n2. Privacidad de Datos: Cumplimos con altos estándares de seguridad médica para proteger su historial clínico.\n3. Puntualidad: Se solicita confirmar o cancelar citas con un mínimo de 12 a 24 horas de antelación.`;
  });
  const [privacyPolicy, setPrivacyPolicy] = useState<string>(() => {
    return localStorage.getItem("dianasrl_privacy_policy") || 
      `POLÍTICAS DE PRIVACIDAD\n\n1. Información Recopilada: Recopilamos datos de contacto, DNI/Cédula/Pasaporte e historial clínico de forma confidencial.\n2. Seguridad: Su información clínica es accesible solo por el personal clínico autorizado.\n3. Solicitudes de Acceso: Puede solicitar una copia de su odontograma o expediente dental contactando directamente a recepción.`;
  });
  const [showPoliciesModal, setShowPoliciesModal] = useState<boolean>(false);
  const [activePoliciesTab, setActivePoliciesTab] = useState<"terms" | "privacy">("terms");

  const t = translations[lang];
  const isDark = theme === "dark";

  // Tailwind style bindings depending on dark vs light mode & chosen brand theme
  const themeColors = {
    pink: {
      accentText: "text-pink-500",
      accentTextLight: "text-pink-300",
      accentBgMutedDark: "bg-pink-950/20",
      accentBgLight10: "bg-pink-500/10",
      accentBorderLight: "border-pink-500/10",
      accentBorderMedium: "border-pink-500/20",
      accentFill: "fill-pink-500/10",
      accentFocusRing: "focus:ring-pink-500/20",
      accentFocusBorder: "focus:border-pink-500",
      gradient: "from-pink-400 to-rose-500",
      gradientHover: "hover:from-pink-500 hover:to-rose-600",
      shadow: "shadow-pink-500/10",
      bgMainDark: "bg-[#181113] text-pink-100",
      bgMainLight: "bg-gradient-to-br from-[#fff6f7] via-white to-[#fffbfb] text-rose-950",
      bgCardDark: "bg-[#251b1d] border border-pink-500/10",
      bgCardLight: "bg-white border border-rose-100/70 shadow-sm shadow-rose-100/20",
      textSubLight: "text-pink-300/70",
      textSubDark: "text-rose-600/80",
      inputStyleDark: "bg-[#1e1416] border border-pink-500/20 text-white focus:border-pink-500 focus:ring-pink-500/20",
      inputStyleLight: "bg-rose-50/20 border border-rose-200 text-rose-950 focus:border-rose-400 focus:ring-rose-500/10",
      badgeColor: "bg-pink-500/10 text-pink-500",
      themeAccentName: "pink-500"
    },
    emerald: {
      accentText: "text-emerald-500",
      accentTextLight: "text-emerald-300",
      accentBgMutedDark: "bg-emerald-950/20",
      accentBgLight10: "bg-emerald-500/10",
      accentBorderLight: "border-emerald-500/10",
      accentBorderMedium: "border-emerald-500/20",
      accentFill: "fill-emerald-500/10",
      accentFocusRing: "focus:ring-emerald-500/20",
      accentFocusBorder: "focus:border-emerald-500",
      gradient: "from-emerald-400 to-teal-500",
      gradientHover: "hover:from-emerald-500 hover:to-teal-600",
      shadow: "shadow-emerald-500/10",
      bgMainDark: "bg-[#111815] text-emerald-100",
      bgMainLight: "bg-gradient-to-br from-[#f0fdf4] via-white to-[#fbfdfb] text-emerald-950",
      bgCardDark: "bg-[#1b2520] border border-emerald-500/10",
      bgCardLight: "bg-white border border-emerald-100/70 shadow-sm shadow-emerald-100/20",
      textSubLight: "text-emerald-300/70",
      textSubDark: "text-emerald-600/80",
      inputStyleDark: "bg-[#141e19] border border-emerald-500/20 text-white focus:border-emerald-500 focus:ring-emerald-500/20",
      inputStyleLight: "bg-emerald-50/20 border border-emerald-200 text-emerald-950 focus:border-emerald-400 focus:ring-emerald-500/10",
      badgeColor: "bg-emerald-500/10 text-emerald-500",
      themeAccentName: "emerald-500"
    },
    amber: {
      accentText: "text-amber-500",
      accentTextLight: "text-amber-300",
      accentBgMutedDark: "bg-amber-950/20",
      accentBgLight10: "bg-amber-500/10",
      accentBorderLight: "border-amber-500/10",
      accentBorderMedium: "border-amber-500/20",
      accentFill: "fill-amber-500/10",
      accentFocusRing: "focus:ring-amber-500/20",
      accentFocusBorder: "focus:border-amber-500",
      gradient: "from-amber-400 to-orange-500",
      gradientHover: "hover:from-amber-500 hover:to-orange-600",
      shadow: "shadow-amber-500/10",
      bgMainDark: "bg-[#181611] text-amber-100",
      bgMainLight: "bg-gradient-to-br from-[#fdfbeb] via-white to-[#fdfdfb] text-amber-950",
      bgCardDark: "bg-[#25221b] border border-amber-500/10",
      bgCardLight: "bg-white border border-amber-100/70 shadow-sm shadow-amber-100/20",
      textSubLight: "text-amber-300/70",
      textSubDark: "text-amber-600/80",
      inputStyleDark: "bg-[#1e1c14] border border-amber-500/20 text-white focus:border-amber-500 focus:ring-amber-500/20",
      inputStyleLight: "bg-amber-50/20 border border-amber-200 text-amber-950 focus:border-amber-400 focus:ring-amber-500/10",
      badgeColor: "bg-amber-500/10 text-amber-600",
      themeAccentName: "amber-500"
    },
    violet: {
      accentText: "text-violet-500",
      accentTextLight: "text-violet-300",
      accentBgMutedDark: "bg-violet-950/20",
      accentBgLight10: "bg-violet-500/10",
      accentBorderLight: "border-violet-500/10",
      accentBorderMedium: "border-violet-500/20",
      accentFill: "fill-violet-500/10",
      accentFocusRing: "focus:ring-violet-500/20",
      accentFocusBorder: "focus:border-violet-500",
      gradient: "from-violet-400 to-purple-500",
      gradientHover: "hover:from-violet-500 hover:to-purple-600",
      shadow: "shadow-violet-500/10",
      bgMainDark: "bg-[#141118] text-violet-100",
      bgMainLight: "bg-gradient-to-br from-[#faf5ff] via-white to-[#fafaff] text-violet-950",
      bgCardDark: "bg-[#1f1b25] border border-violet-500/10",
      bgCardLight: "bg-white border border-violet-100/70 shadow-sm shadow-violet-100/20",
      textSubLight: "text-violet-300/70",
      textSubDark: "text-violet-600/80",
      inputStyleDark: "bg-[#19141e] border border-violet-500/20 text-white focus:border-violet-500 focus:ring-violet-500/20",
      inputStyleLight: "bg-violet-50/20 border border-violet-200 text-violet-950 focus:border-violet-400 focus:ring-violet-500/10",
      badgeColor: "bg-violet-500/10 text-violet-500",
      themeAccentName: "violet-500"
    },
    sky: {
      accentText: "text-sky-500",
      accentTextLight: "text-sky-300",
      accentBgMutedDark: "bg-sky-950/20",
      accentBgLight10: "bg-sky-500/10",
      accentBorderLight: "border-sky-500/10",
      accentBorderMedium: "border-sky-500/20",
      accentFill: "fill-sky-500/10",
      accentFocusRing: "focus:ring-sky-500/20",
      accentFocusBorder: "focus:border-sky-500",
      gradient: "from-sky-400 to-blue-500",
      gradientHover: "hover:from-sky-500 hover:to-blue-600",
      shadow: "shadow-sky-500/10",
      bgMainDark: "bg-[#111518] text-sky-100",
      bgMainLight: "bg-gradient-to-br from-[#f0f9ff] via-white to-[#fbfdff] text-sky-950",
      bgCardDark: "bg-[#1b2125] border border-sky-500/10",
      bgCardLight: "bg-white border border-sky-100/70 shadow-sm shadow-sky-100/20",
      textSubLight: "text-sky-300/70",
      textSubDark: "text-sky-600/80",
      inputStyleDark: "bg-[#141a1e] border border-sky-500/20 text-white focus:border-sky-500 focus:ring-sky-500/20",
      inputStyleLight: "bg-sky-50/20 border border-sky-200 text-sky-950 focus:border-sky-400 focus:ring-sky-500/10",
      badgeColor: "bg-sky-500/10 text-sky-500",
      themeAccentName: "sky-500"
    }
  };

  const activeThemeObj = themeColors[clinicColorTheme] || themeColors.pink;

  const bgMain = isDark 
    ? `${activeThemeObj.bgMainDark} select-none transition-colors duration-300` 
    : `${activeThemeObj.bgMainLight} select-none transition-colors duration-300`;
  
  const bgCard = isDark 
    ? `${activeThemeObj.bgCardDark} shadow-lg shadow-black/40 rounded-3xl p-6` 
    : `${activeThemeObj.bgCardLight} rounded-3xl p-6`;

  const textTitle = isDark ? "text-white font-display" : "text-gray-900 font-display";
  const textSub = isDark ? activeThemeObj.textSubLight : activeThemeObj.textSubDark;
  const borderCol = isDark ? activeThemeObj.accentBorderLight : "border-rose-100/80";
  
  const inputStyle = isDark 
    ? `w-full px-4 py-2.5 rounded-xl focus:outline-none focus:ring-2 text-sm transition-all ${activeThemeObj.inputStyleDark}` 
    : `w-full px-4 py-2.5 rounded-xl focus:outline-none focus:ring-2 text-sm transition-all ${activeThemeObj.inputStyleLight}`;

  // Synchronize document/tab title with the current clinic name and portal mode
  useEffect(() => {
    if (portalMode === "patient") {
      document.title = `${clinicName} — ${lang === "es" ? "Portal de Pacientes" : "Patient Portal"}`;
    } else {
      document.title = `${clinicName} — ${lang === "es" ? "Consola Dental" : "Dental Console"}`;
    }
  }, [clinicName, portalMode, lang]);

  // Initialize and Seed mock database values if empty
  useEffect(() => {
    // 1. Language and Theme
    const savedLang = localStorage.getItem("dianasrl_lang");
    if (savedLang === "en" || savedLang === "es" || savedLang === "fr" || savedLang === "pt") {
      setLang(savedLang);
    }
    const savedTheme = localStorage.getItem("dianasrl_theme");
    if (savedTheme === "dark" || savedTheme === "light") {
      setTheme(savedTheme as any);
    }

    // 2. Patients Seed (Purge mock data if it exists in local storage to keep only Google Sheets data)
    let localPatients = localStorage.getItem("dianasrl_patients");
    if (localPatients) {
      try {
        const parsed: Patient[] = JSON.parse(localPatients);
        const hasMock = parsed.some(
          p =>
            p.id === "p1" ||
            p.id === "p2" ||
            p.id === "p3" ||
            p.name.includes("Diana María Pérez") ||
            p.name.includes("Alexander Rodríguez") ||
            p.name.includes("Camila Santos Ortiz")
        );
        if (hasMock) {
          localStorage.removeItem("dianasrl_patients");
          localStorage.removeItem("dianasrl_appointments");
          localStorage.removeItem("dianasrl_odontograms");
          localStorage.removeItem("dianasrl_payments");
          localPatients = null;
        }
      } catch (e) {
        localPatients = null;
      }
    }

    let loadedPatients: Patient[] = [];
    if (localPatients) {
      try {
        loadedPatients = JSON.parse(localPatients);
      } catch (e) {
        loadedPatients = [];
      }
      setPatients(loadedPatients);
    } else {
      setPatients([]);
      localStorage.setItem("dianasrl_patients", JSON.stringify([]));
    }
    if (loadedPatients.length > 0) {
      setSelectedPatientId(loadedPatients[0].id);
    } else {
      setSelectedPatientId("");
    }

    // 3. Appointments Seed
    const localAppointments = localStorage.getItem("dianasrl_appointments");
    if (localAppointments && localPatients) { // Only load appointments if we kept the patient list
      try {
        setAppointments(JSON.parse(localAppointments));
      } catch (e) {
        setAppointments([]);
      }
    } else {
      setAppointments([]);
      localStorage.setItem("dianasrl_appointments", JSON.stringify([]));
    }

    // 4. Odontogram Seed
    const localOdontograms = localStorage.getItem("dianasrl_odontograms");
    if (localOdontograms && localPatients) {
      try {
        setToothConditions(JSON.parse(localOdontograms));
      } catch (e) {
        setToothConditions({});
      }
    } else {
      setToothConditions({});
      localStorage.setItem("dianasrl_odontograms", JSON.stringify({}));
    }

    // 5. Payments Seed
    const localPayments = localStorage.getItem("dianasrl_payments");
    if (localPayments && localPatients) {
      try {
        setPayments(JSON.parse(localPayments));
      } catch (e) {
        setPayments([]);
      }
    } else {
      setPayments([]);
      localStorage.setItem("dianasrl_payments", JSON.stringify([]));
    }

    // 6. Sync Logs
    const localLogs = localStorage.getItem("dianasrl_sync_logs");
    if (localLogs) {
      setSyncLogs(JSON.parse(localLogs));
    } else {
      const seedLogs: SyncLog[] = [
        { id: "l1", timestamp: new Date(Date.now() - 3600000).toISOString(), action: "Database Initialized", status: "Success", message: "Conexión local establecida y semilla de datos cargada." }
      ];
      setSyncLogs(seedLogs);
      localStorage.setItem("dianasrl_sync_logs", JSON.stringify(seedLogs));
    }

    // 6.5. Doctor Availabilities Seed
    const localAvailabilities = localStorage.getItem("dianasrl_availabilities");
    if (localAvailabilities) {
      setAvailabilities(JSON.parse(localAvailabilities));
    } else {
      const getRelativeDate = (offsetDays: number) => {
        const d = new Date();
        d.setDate(d.getDate() + offsetDays);
        return d.toISOString().split("T")[0];
      };
      const seedAvails: DoctorAvailability[] = [
        {
          id: "av1",
          doctorName: "Dra. Diana Rojas",
          date: getRelativeDate(0),
          slots: ["08:00", "09:00", "10:00", "11:00", "14:00", "15:00", "16:00"]
        },
        {
          id: "av2",
          doctorName: "Dra. Diana Rojas",
          date: getRelativeDate(1),
          slots: ["09:00", "10:00", "11:00", "14:00", "15:00"]
        },
        {
          id: "av3",
          doctorName: "Dr. Marcos Soler",
          date: getRelativeDate(1),
          slots: ["08:00", "10:00", "14:00", "16:00"]
        },
        {
          id: "av4",
          doctorName: "Dra. Lucía Santos",
          date: getRelativeDate(2),
          slots: ["09:00", "10:00", "11:00", "15:00", "16:00"]
        }
      ];
      setAvailabilities(seedAvails);
      localStorage.setItem("dianasrl_availabilities", JSON.stringify(seedAvails));
    }

    // 7. Admin Accounts Seeding (Master Founder Account and any created admins with real-time credentials)
    const localAdmins = localStorage.getItem("dianasrl_admins");
    let initialAdmins: AdminUser[] = [];
    if (localAdmins) {
      try {
        initialAdmins = JSON.parse(localAdmins);
      } catch (e) {
        initialAdmins = [];
      }
    }

    // Always enforce Master founder credentials exactly
    const masterIndex = initialAdmins.findIndex(a => a.id === "adm_master" || a.role === "master");
    const masterAccount: AdminUser = {
      id: "adm_master",
      username: "Dianasrl@#",
      name: "Dra. Diana Rojas",
      passwordHash: "FranciscoEmmanuelYDianaRoja2026@#",
      role: "master",
      createdAt: new Date().toISOString()
    };

    if (masterIndex !== -1) {
      initialAdmins[masterIndex] = masterAccount;
    } else {
      initialAdmins.unshift(masterAccount);
    }

    setAdmins(initialAdmins);
    localStorage.setItem("dianasrl_admins", JSON.stringify(initialAdmins));

    // Also update any active admin session with correct credentials so they stay synchronized
    const savedSession = localStorage.getItem("dianasrl_current_admin");
    if (savedSession) {
      try {
        const parsed = JSON.parse(savedSession);
        if (parsed.role === "master") {
          localStorage.setItem("dianasrl_current_admin", JSON.stringify(masterAccount));
          setCurrentAdmin(masterAccount);
        }
      } catch (e) {
        // ignore
      }
    }
  }, []);

  // Synchronize availDoctor with logged-in admin
  useEffect(() => {
    if (currentAdmin) {
      setAvailDoctor(currentAdmin.name);
    } else {
      setAvailDoctor("Dra. Diana Rojas");
    }
  }, [currentAdmin]);

  // Show customized toast message helper
  const triggerToast = (text: string, type: "success" | "error" = "success") => {
    setToastMessage({ text, type });
    setTimeout(() => setToastMessage(null), 4000);
  };

  // Switch Theme
  const toggleTheme = () => {
    const nextTheme = theme === "light" ? "dark" : "light";
    setTheme(nextTheme);
    localStorage.setItem("dianasrl_theme", nextTheme);
  };

  // Switch Language
  const changeLanguage = (next: "en" | "es" | "fr" | "pt") => {
    setLang(next);
    localStorage.setItem("dianasrl_lang", next);
  };

  // Sync API fetch proxy to write real rows into Google Spreadsheet
  const executeSync = async (action: string, payloadData: any) => {
    const logId = "l_" + Math.random().toString(36).substring(2, 9);
    const newLogEntry: SyncLog = {
      id: logId,
      timestamp: new Date().toISOString(),
      action: `POST to ${action}`,
      status: "Pending",
      message: `Enviando registro a Google Sheets...`
    };

    // Update state & persist logs locally
    const currentLogs = [newLogEntry, ...syncLogs].slice(0, 50);
    setSyncLogs(currentLogs);
    localStorage.setItem("dianasrl_sync_logs", JSON.stringify(currentLogs));

    if (!liveSync || !scriptUrl) {
      newLogEntry.status = "Error";
      newLogEntry.message = "Sincronización deshabilitada o URL vacía. Registro guardado localmente.";
      const updatedLogs = currentLogs.map(l => l.id === logId ? newLogEntry : l);
      setSyncLogs(updatedLogs);
      localStorage.setItem("dianasrl_sync_logs", JSON.stringify(updatedLogs));
      return;
    }

    try {
      const response = await fetch("/api/sync", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          url: scriptUrl,
          payload: {
            action: action,
            timestamp: new Date().toISOString(),
            data: payloadData,
            spreadsheetUrl: spreadsheetUrl
          }
        })
      });

      if (response.status === 404) {
        throw new Error("El servidor proxy no se encuentra activo o la ruta de sincronización es incorrecta.");
      }

      const contentType = response.headers.get("content-type");
      if (contentType && contentType.includes("text/html")) {
        throw new Error("El servidor proxy o Google Sheets retornó una página HTML en lugar de JSON. Asegúrate de que tu Web App de Google Apps Script esté publicada para 'Cualquiera' (Anyone) y autorizada.");
      }

      const resJson = await response.json();
      if (resJson.success) {
        newLogEntry.status = "Success";
        newLogEntry.message = `Sincronizado con Sheets con éxito. Fila agregada en tabla: ${action}.`;
        triggerToast(`${t.syncSuccess} (${action})`);
      } else {
        newLogEntry.status = "Error";
        newLogEntry.message = `Apps Script respondió con error: ${resJson.error || "Código de error no documentado."}`;
        triggerToast(t.syncError, "error");
      }
    } catch (err: any) {
      newLogEntry.status = "Error";
      newLogEntry.message = `Fallo de conexión de red: ${err.message || err}. Registro guardado localmente en caché segura.`;
      triggerToast(t.syncError, "error");
    }

    const finalLogs = currentLogs.map(l => l.id === logId ? { ...newLogEntry } : l);
    setSyncLogs(finalLogs);
    localStorage.setItem("dianasrl_sync_logs", JSON.stringify(finalLogs));
  };

  // Test sync connection ping
  const handleTestSync = async () => {
    setIsTestingSync(true);
    const logId = "l_" + Math.random().toString(36).substring(2, 9);
    const newLogEntry: SyncLog = {
      id: logId,
      timestamp: new Date().toISOString(),
      action: "Test Ping Connection",
      status: "Pending",
      message: t.testing
    };

    setSyncLogs(prev => [newLogEntry, ...prev]);

    try {
      const response = await fetch("/api/sync", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          url: scriptUrl,
          payload: { 
            action: "ping", 
            data: { client: "DianaSRL App", status: "checking" },
            spreadsheetUrl: spreadsheetUrl
          }
        })
      });

      if (response.status === 404) {
        throw new Error("El servidor proxy no se encuentra activo o la ruta de sincronización es incorrecta.");
      }

      const contentType = response.headers.get("content-type");
      if (contentType && contentType.includes("text/html")) {
        throw new Error("El servidor proxy o Google Sheets retornó una página HTML en lugar de JSON. Asegúrate de que tu Web App de Google Apps Script esté publicada para 'Cualquiera' (Anyone) y autorizada.");
      }

      const resJson = await response.json();
      if (resJson.success) {
        newLogEntry.status = "Success";
        newLogEntry.message = "Ping exitoso. El script de Google Sheets respondió de forma activa.";
        triggerToast("¡Conexión de Google Sheets verificada con éxito!");
      } else {
        newLogEntry.status = "Error";
        newLogEntry.message = `Ping fallido: ${resJson.error || "El script no respondió con estado exitoso."}`;
        triggerToast("Fallo en la prueba de conexión.", "error");
      }
    } catch (err: any) {
      newLogEntry.status = "Error";
      newLogEntry.message = `Error de red al conectar al Web App: ${err.message || err}`;
      triggerToast("Fallo de red al conectar.", "error");
    }

    setSyncLogs(prev => prev.map(l => l.id === logId ? { ...newLogEntry } : l));
    setIsTestingSync(false);
  };

  // Helper to extract values from Google Sheets rows in a case-insensitive and multi-language way
  const getFieldValue = (obj: any, keys: string[], defaultValue: any = ""): any => {
    if (!obj || typeof obj !== "object") return defaultValue;
    const objKeys = Object.keys(obj);
    
    // Try exact matches first
    for (const key of keys) {
      if (obj[key] !== undefined && obj[key] !== null && obj[key] !== "") {
        return obj[key];
      }
    }
    
    // Try case-insensitive and whitespace-stripped matches
    const cleanedKeys = keys.map(k => k.toLowerCase().replace(/\s+/g, ""));
    for (const objKey of objKeys) {
      const cleanedObjKey = objKey.toLowerCase().replace(/\s+/g, "");
      if (cleanedKeys.includes(cleanedObjKey)) {
        if (obj[objKey] !== undefined && obj[objKey] !== null && obj[objKey] !== "") {
          return obj[objKey];
        }
      }
    }
    
    return defaultValue;
  };

  // Import and Synchronize All Data from Google Sheets
  const handleImportAllFromSheets = async () => {
    setIsImportingSync(true);
    const logId = "l_" + Math.random().toString(36).substring(2, 9);
    const newLogEntry: SyncLog = {
      id: logId,
      timestamp: new Date().toISOString(),
      action: "Import All Data",
      status: "Pending",
      message: "Consultando datos de la base de datos de Google Sheets..."
    };

    setSyncLogs(prev => [newLogEntry, ...prev]);

    try {
      const response = await fetch("/api/sync", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          url: scriptUrl,
          payload: { 
            action: "read_all", 
            data: { client: "DianaSRL App" },
            spreadsheetUrl: spreadsheetUrl
          }
        })
      });

      if (response.status === 404) {
        throw new Error("El servidor proxy no se encuentra activo o la ruta de sincronización es incorrecta.");
      }

      const contentType = response.headers.get("content-type");
      if (contentType && contentType.includes("text/html")) {
        throw new Error("El servidor proxy o Google Sheets retornó una página HTML en lugar de JSON. Asegúrate de que tu Web App de Google Apps Script esté publicada para 'Cualquiera' (Anyone) y autorizada.");
      }

      const resJson = await response.json();
      
      if (resJson.success && resJson.data) {
        const sheetData = resJson.data;
        let importedCount = 0;

        // Parse patients
        const rawPatients = sheetData.patients || sheetData.Pacientes || sheetData.pacientes;
        if (rawPatients && Array.isArray(rawPatients)) {
          const parsed = rawPatients.map((p: any) => {
            const rawId = getFieldValue(p, ["id", "ID", "PatientID", "patientId", "PacienteID", "pacienteId", "idPaciente"]);
            return {
              id: rawId ? String(rawId) : ("p_" + Math.random().toString(36).substring(2, 9)),
              name: String(getFieldValue(p, ["name", "Nombre", "Name", "patientName", "pacienteNombre"])),
              phone: String(getFieldValue(p, ["phone", "Teléfono", "telefono", "Phone", "Celular", "celular"])),
              email: String(getFieldValue(p, ["email", "Correo", "correo", "Email"])),
              dni: String(getFieldValue(p, ["dni", "DNI", "Cédula", "cedula", "Dni"])),
              dob: String(getFieldValue(p, ["dob", "DOB", "FechaNacimiento", "Fecha de Nacimiento", "fechaNacimiento", "nacimiento"])),
              notes: String(getFieldValue(p, ["notes", "Notas", "notas", "Notes", "Observaciones", "observaciones"])),
              registeredAt: String(getFieldValue(p, ["registeredAt", "FechaRegistro", "Fecha Registro", "registrado", "registered"], new Date().toISOString()))
            };
          });
          if (parsed.length > 0) {
            setPatients(parsed);
            localStorage.setItem("dianasrl_patients", JSON.stringify(parsed));
            setSelectedPatientId(parsed[0].id);
            importedCount += parsed.length;
          }
        }

        // Parse appointments
        const rawAppointments = sheetData.appointments || sheetData.Citas || sheetData.citas;
        if (rawAppointments && Array.isArray(rawAppointments)) {
          const parsed = rawAppointments.map((a: any) => {
            const rawId = getFieldValue(a, ["id", "ID", "AppointmentId", "CitaID", "citaId"]);
            return {
              id: rawId ? String(rawId) : ("a_" + Math.random().toString(36).substring(2, 9)),
              patientId: String(getFieldValue(a, ["patientId", "PacienteID", "patient_id", "pacienteId", "idPaciente"])),
              patientName: String(getFieldValue(a, ["patientName", "Paciente", "Nombre Paciente", "nombrePaciente", "patient_name"])),
              dentistName: String(getFieldValue(a, ["dentistName", "Dentista", "Odontólogo", "odontologo", "dentist_name", "doctor"])),
              date: String(getFieldValue(a, ["date", "Fecha", "fecha", "Date"])),
              time: String(getFieldValue(a, ["time", "Hora", "hora", "Time"])),
              treatmentType: String(getFieldValue(a, ["treatmentType", "Tratamiento", "TipoTratamiento", "treatment_type", "tratamiento"])),
              notes: String(getFieldValue(a, ["notes", "Notas", "notas", "Notes", "Observaciones", "observaciones"])),
              status: String(getFieldValue(a, ["status", "Estado", "estado", "Status"], "Scheduled"))
            };
          });
          if (parsed.length > 0) {
            setAppointments(parsed);
            localStorage.setItem("dianasrl_appointments", JSON.stringify(parsed));
            importedCount += parsed.length;
          }
        }

        // Parse payments
        const rawPayments = sheetData.payments || sheetData.Pagos || sheetData.pagos;
        if (rawPayments && Array.isArray(rawPayments)) {
          const parsed = rawPayments.map((p: any) => {
            const rawId = getFieldValue(p, ["id", "ID", "PaymentId", "PagoID", "pagoId"]);
            return {
              id: rawId ? String(rawId) : ("pay_" + Math.random().toString(36).substring(2, 9)),
              patientId: String(getFieldValue(p, ["patientId", "PacienteID", "patient_id", "pacienteId", "idPaciente"])),
              patientName: String(getFieldValue(p, ["patientName", "Paciente", "Nombre Paciente", "nombrePaciente", "patient_name"])),
              treatmentType: String(getFieldValue(p, ["treatmentType", "Tratamiento", "TipoTratamiento", "treatment_type", "tratamiento"])),
              amountTotal: Number(getFieldValue(p, ["amountTotal", "MontoTotal", "Total", "monto_total", "montoTotal"])) || 0,
              amountPaid: Number(getFieldValue(p, ["amountPaid", "MontoPagado", "Pagado", "monto_pagado", "montoPagado"])) || 0,
              amountPending: Number(getFieldValue(p, ["amountPending", "MontoPendiente", "Pendiente", "monto_pendiente", "montoPending"])) || 0,
              paymentMethod: String(getFieldValue(p, ["paymentMethod", "Método", "MetodoPago", "metodo_pago", "payment_method", "metodo"], "Cash")),
              paymentStatus: String(getFieldValue(p, ["paymentStatus", "EstadoPago", "Estado Pago", "estado_pago", "status"], "Paid")),
              date: String(getFieldValue(p, ["date", "Fecha", "fecha", "Date"]))
            };
          });
          if (parsed.length > 0) {
            setPayments(parsed);
            localStorage.setItem("dianasrl_payments", JSON.stringify(parsed));
            importedCount += parsed.length;
          }
        }

        // Parse odontograms
        const rawOdontograms = sheetData.odontograms || sheetData.odontogramas || sheetData.odontograma || sheetData.Odontograma || sheetData.Odontogramas;
        if (rawOdontograms) {
          let parsed: Record<string, ToothCondition[]> = {};
          if (Array.isArray(rawOdontograms)) {
            rawOdontograms.forEach((o: any) => {
              const pid = getFieldValue(o, ["patientId", "PacienteID", "patient_id", "pacienteId", "id"]);
              if (pid) {
                const patId = String(pid);
                if (!parsed[patId]) parsed[patId] = [];
                parsed[patId].push({
                  toothNumber: Number(getFieldValue(o, ["toothNumber", "Diente", "Pieza", "tooth_number", "diente", "pieza"])) || 11,
                  status: String(getFieldValue(o, ["status", "Estado", "Condición", "condicion", "estado", "status"], "Healthy")) as any,
                  notes: String(getFieldValue(o, ["notes", "Notas", "notas", "Observaciones", "observaciones"])),
                  updatedAt: String(getFieldValue(o, ["updatedAt", "Fecha", "fecha", "updated_at"], new Date().toISOString()))
                });
              }
            });
          } else if (typeof rawOdontograms === "object") {
            parsed = rawOdontograms;
          }
          setToothConditions(parsed);
          localStorage.setItem("dianasrl_odontograms", JSON.stringify(parsed));
        }

        newLogEntry.status = "Success";
        newLogEntry.message = `Sincronización bidireccional exitosa. Se cargaron e integraron ${importedCount} registros clínicos de Google Sheets.`;
        triggerToast(`¡Base de datos cargada! ${importedCount} registros clínicos importados con éxito.`);
      } else {
        newLogEntry.status = "Error";
        newLogEntry.message = "La Web App de Sheets no retornó un arreglo de datos estructurado. Asegure que su Apps Script soporte 'read_all'.";
        triggerToast("La Web App no retornó datos válidos.", "error");
      }
    } catch (err: any) {
      newLogEntry.status = "Error";
      newLogEntry.message = `Fallo de conexión o de formato: ${err.message || err}`;
      triggerToast("Fallo al conectar con Google Sheets.", "error");
    }

    setSyncLogs(prev => prev.map(l => l.id === logId ? { ...newLogEntry } : l));
    setIsImportingSync(false);
  };

  // Export and Mirror All Current Local Data to Google Sheets
  const handleExportAllToSheets = async () => {
    setIsExportingSync(true);
    const logId = "l_" + Math.random().toString(36).substring(2, 9);
    const newLogEntry: SyncLog = {
      id: logId,
      timestamp: new Date().toISOString(),
      action: "Export All Data",
      status: "Pending",
      message: "Enviando todos los registros locales a Google Sheets..."
    };

    setSyncLogs(prev => [newLogEntry, ...prev]);

    // Flatten odontograms so they can be written as flat rows in Google Sheets
    const flatOdontograms: any[] = [];
    Object.entries(toothConditions).forEach(([patId, conditions]) => {
      (conditions as ToothCondition[]).forEach(cond => {
        flatOdontograms.push({
          patientId: patId,
          toothNumber: cond.toothNumber,
          status: cond.status,
          notes: cond.notes,
          updatedAt: cond.updatedAt
        });
      });
    });

    const payload = {
      action: "bulk_write",
      data: {
        Pacientes: patients,
        Citas: appointments,
        Pagos: payments,
        Odontogramas: flatOdontograms
      }
    };

    try {
      const response = await fetch("/api/sync", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          url: scriptUrl,
          payload: payload
        })
      });

      if (response.status === 404) {
        throw new Error("El servidor proxy no se encuentra activo o la ruta de sincronización es incorrecta.");
      }

      const contentType = response.headers.get("content-type");
      if (contentType && contentType.includes("text/html")) {
        throw new Error("El servidor proxy o Google Sheets retornó una página HTML en lugar de JSON. Asegúrate de que tu Web App de Google Apps Script esté publicada para 'Cualquiera' (Anyone) y autorizada.");
      }

      const resJson = await response.json();
      
      if (resJson.success) {
        newLogEntry.status = "Success";
        newLogEntry.message = `Sincronización de subida completada con éxito. Se enviaron ${patients.length} pacientes, ${appointments.length} citas y ${payments.length} pagos.`;
        triggerToast("¡Sincronización completada! Todos los datos locales se han subido a Google Sheets.");
      } else {
        newLogEntry.status = "Error";
        newLogEntry.message = `El script de Google Sheets retornó un error: ${resJson.error || "Fallo desconocido"}.`;
        triggerToast("Error al subir los datos a Google Sheets.", "error");
      }
    } catch (err: any) {
      newLogEntry.status = "Error";
      newLogEntry.message = `Fallo al conectar con el servidor o script: ${err.message || err}`;
      triggerToast("Fallo al conectar con Google Sheets.", "error");
    }

    setSyncLogs(prev => prev.map(l => l.id === logId ? { ...newLogEntry } : l));
    setIsExportingSync(false);
  };

  // Register Patient Action
  const handleRegisterPatient = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newPatient.name || !newPatient.dni) {
      triggerToast("El nombre y la cédula son requeridos para registrar un paciente.", "error");
      return;
    }

    const created: Patient = {
      id: "p_" + Math.random().toString(36).substring(2, 9),
      name: newPatient.name,
      phone: newPatient.phone || "",
      email: newPatient.email || "",
      dni: newPatient.dni,
      dob: newPatient.dob || "",
      notes: newPatient.notes || "",
      registeredAt: new Date().toISOString()
    };

    const updatedList = [created, ...patients];
    setPatients(updatedList);
    localStorage.setItem("dianasrl_patients", JSON.stringify(updatedList));
    setSelectedPatientId(created.id);

    // Trigger Sheet sync
    executeSync("Pacientes", created);

    // Reset Form
    setNewPatient({ name: "", phone: "", email: "", dni: "", dob: "", notes: "" });
    triggerToast(`Paciente ${created.name} registrado con éxito.`);
  };

  // Schedule Appointment Action
  const handleScheduleAppointment = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newAppointment.patientId || !newAppointment.date || !newAppointment.time || !newAppointment.treatmentType) {
      triggerToast("Todos los campos con asterisco son requeridos.", "error");
      return;
    }

    const patient = patients.find(p => p.id === newAppointment.patientId);
    if (!patient) return;

    const created: Appointment = {
      id: "a_" + Math.random().toString(36).substring(2, 9),
      patientId: newAppointment.patientId,
      patientName: patient.name,
      dentistName: newAppointment.dentistName || "Dra. Diana Rojas",
      date: newAppointment.date,
      time: newAppointment.time,
      treatmentType: newAppointment.treatmentType,
      notes: newAppointment.notes || "",
      status: "Scheduled"
    };

    const updatedList = [created, ...appointments];
    setAppointments(updatedList);
    localStorage.setItem("dianasrl_appointments", JSON.stringify(updatedList));

    // Trigger Sheet sync
    executeSync("Citas", created);

    // Reset Form
    setNewAppointment({ patientId: "", dentistName: "", date: "", time: "", treatmentType: "", notes: "" });
    triggerToast("Cita médica agendada correctamente.");
  };

  // Update Appointment Status Action
  const handleUpdateAppointmentStatus = (appId: string, nextStatus: "Scheduled" | "Completed" | "Cancelled") => {
    const updated = appointments.map(app => {
      if (app.id === appId) {
        const item = { ...app, status: nextStatus };
        // Trigger Sheets sync on update
        executeSync("ActualizarCita", item);
        return item;
      }
      return app;
    });
    setAppointments(updated);
    localStorage.setItem("dianasrl_appointments", JSON.stringify(updated));
    triggerToast(`Cita actualizada a: ${nextStatus === "Completed" ? t.statusCompleted : nextStatus === "Cancelled" ? t.statusCancelled : t.statusScheduled}`);
  };

  // Apply Tooth Condition Action
  const handleApplyToothCondition = (status: "Healthy" | "Caries" | "Endo" | "Missing" | "Crown" | "Filling", notes = "") => {
    if (!selectedPatientId || selectedTooth === null) return;

    const patientConditions = toothConditions[selectedPatientId] || [];
    const conditionIndex = patientConditions.findIndex(c => c.toothNumber === selectedTooth);

    const newCondition: ToothCondition = {
      toothNumber: selectedTooth,
      status,
      notes: notes || undefined,
      updatedAt: new Date().toISOString()
    };

    let updatedPatientConditions = [...patientConditions];
    if (conditionIndex >= 0) {
      updatedPatientConditions[conditionIndex] = newCondition;
    } else {
      updatedPatientConditions.push(newCondition);
    }

    const nextToothConditions = {
      ...toothConditions,
      [selectedPatientId]: updatedPatientConditions
    };

    setToothConditions(nextToothConditions);
    localStorage.setItem("dianasrl_odontograms", JSON.stringify(nextToothConditions));

    // Sync with Sheets
    executeSync("Odontogramas", {
      patientId: selectedPatientId,
      patientName: patients.find(p => p.id === selectedPatientId)?.name || "Paciente",
      toothNumber: selectedTooth,
      status,
      notes: notes || "Ninguna"
    });

    setSelectedTooth(null);
    triggerToast(`Diente ${selectedTooth} actualizado a: ${status}`);
  };

  // Register Billing Payment Action
  const handleRegisterPayment = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newPayment.patientId || !newPayment.treatmentType || !newPayment.amountTotal) {
      triggerToast("Por favor complete los datos obligatorios del cobro.", "error");
      return;
    }

    const patient = patients.find(p => p.id === newPayment.patientId);
    if (!patient) return;

    const total = Number(newPayment.amountTotal);
    const paid = Number(newPayment.amountPaid || 0);
    const pending = Math.max(0, total - paid);
    const status: "Pending" | "Partial" | "Paid" = paid === 0 ? "Pending" : pending === 0 ? "Paid" : "Partial";

    const created: PaymentRecord = {
      id: "pay_" + Math.random().toString(36).substring(2, 9),
      patientId: newPayment.patientId,
      patientName: patient.name,
      treatmentType: newPayment.treatmentType,
      amountTotal: total,
      amountPaid: paid,
      amountPending: pending,
      paymentMethod: (newPayment.paymentMethod as any) || "Cash",
      paymentStatus: status,
      date: new Date().toISOString().split("T")[0]
    };

    const updatedList = [created, ...payments];
    setPayments(updatedList);
    localStorage.setItem("dianasrl_payments", JSON.stringify(updatedList));

    // Trigger Sheet sync
    executeSync("Pagos", created);

    // Reset Form
    setNewPayment({ patientId: "", treatmentType: "", amountTotal: 0, amountPaid: 0, paymentMethod: "Cash" });
    triggerToast("Transacción de pago registrada correctamente.");
  };

  // Gateway configuration helpers
  const handleAddGateway = () => {
    const newId = "gw_" + Math.random().toString(36).substring(2, 9);
    const newGw: PaymentMethodConfig = {
      id: newId,
      name: lang === "es" ? "Nuevo Banco o Método" : "New Payment Method",
      enabled: true,
      instructions: lang === "es" ? "Escriba aquí los detalles de la cuenta o instrucciones de pago para el paciente." : "Write down the account details or payment instructions for the patient.",
      link: "https://"
    };
    const updated = [...gateways, newGw];
    setGateways(updated);
    localStorage.setItem("dianasrl_gateways", JSON.stringify(updated));
    triggerToast(lang === "es" ? "Método de pago añadido." : "Payment method added.", "success");
  };

  const handleUpdateGateway = (id: string, fields: Partial<PaymentMethodConfig>) => {
    const updated = gateways.map(g => g.id === id ? { ...g, ...fields } : g);
    setGateways(updated);
    localStorage.setItem("dianasrl_gateways", JSON.stringify(updated));
  };

  const handleDeleteGateway = (id: string) => {
    const updated = gateways.filter(g => g.id !== id);
    setGateways(updated);
    localStorage.setItem("dianasrl_gateways", JSON.stringify(updated));
    triggerToast(lang === "es" ? "Método de pago eliminado." : "Payment method removed.", "success");
  };

  // Clinical Diagnosis Generator (Gemini Integration)
  const handleAiAnalyze = async () => {
    if (!aiPrompt.trim()) return;
    setIsAiLoading(true);
    setAiResponse("");

    try {
      const response = await fetch("/api/dental-advisor", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ symptoms: aiPrompt, lang: lang })
      });
      const data = await response.json();
      if (data.success) {
        setAiResponse(data.analysis);
      } else {
        setAiResponse("Fallo al contactar el servicio de Inteligencia Artificial.");
      }
    } catch (e) {
      setAiResponse("Ocurrió un error al contactar al servidor clínico.");
    } finally {
      setIsAiLoading(false);
    }
  };

  // Helper values for dashboard stats
  const totalInvoicedAmt = payments.reduce((acc, curr) => acc + curr.amountTotal, 0);
  const totalCollectedAmt = payments.reduce((acc, curr) => acc + curr.amountPaid, 0);
  const pendingReceivableAmt = payments.reduce((acc, curr) => acc + curr.amountPending, 0);
  const scheduledCount = appointments.filter(a => a.status === "Scheduled").length;
  const completedCount = appointments.filter(a => a.status === "Completed").length;

  return (
    <div className={`min-h-screen ${bgMain} font-sans pb-16`}>
      {/* Toast Notification */}
      <AnimatePresence>
        {toastMessage && (
          <motion.div
            initial={{ opacity: 0, y: -50, scale: 0.9 }}
            animate={{ opacity: 1, y: 0, scale: 1 }}
            exit={{ opacity: 0, y: -20, scale: 0.95 }}
            className="fixed top-6 left-1/2 -translate-x-1/2 z-50 flex items-center space-x-3 px-5 py-3 rounded-2xl shadow-xl bg-rose-500 text-white font-medium text-sm"
          >
            {toastMessage.type === "success" ? <CheckCircle size={18} /> : <XCircle size={18} />}
            <span>{toastMessage.text}</span>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Header Navigation */}
      <header className={`sticky top-0 z-40 backdrop-blur-md border-b ${borderCol} ${isDark ? "bg-[#181113]/85" : "bg-white/85"} transition-all`}>
        <div className="max-w-7xl mx-auto px-6 py-4 flex flex-col md:flex-row items-center justify-between gap-4">
          
          {/* Logo Brand */}
          <div className="flex items-center space-x-3 cursor-pointer" onClick={() => {
            if (portalMode === "admin") {
              setActiveTab("agenda");
            } else {
              setPatientActiveTab("my-appointments");
            }
          }}>
            <div className={`w-10 h-10 rounded-2xl overflow-hidden flex items-center justify-center shadow-md bg-white border ${borderCol} shrink-0`}>
              <img 
                src={clinicLogo} 
                alt={`${clinicName} Logo`} 
                className="w-full h-full object-cover"
                referrerPolicy="no-referrer"
              />
            </div>
            <div>
              <h1 className={`text-2xl font-bold tracking-tighter ${textTitle} flex items-center`}>
                {clinicName}
              </h1>
              <p className={`text-[10px] tracking-widest font-mono font-bold ${activeThemeObj.accentText} uppercase`}>{clinicSlogan}</p>
            </div>
          </div>

          {/* Portal Switcher (Clinical Admin vs Patient Portal) */}
          <div className="flex bg-rose-100/40 dark:bg-pink-900/15 p-1 rounded-2xl border border-pink-500/10 select-none font-bold text-[10px] uppercase tracking-wider gap-0.5">
            <button
              type="button"
              onClick={() => {
                changePortalMode("admin");
              }}
              className={`px-3 py-2 rounded-xl transition-all duration-300 flex items-center space-x-1.5 ${
                portalMode === "admin"
                  ? "bg-gradient-to-r from-pink-400 to-rose-500 text-white shadow-sm shadow-pink-500/20"
                  : "text-rose-800 dark:text-pink-300 hover:text-rose-950 dark:hover:text-white"
              }`}
            >
              <Users size={12} />
              <span>{lang === "es" ? "Clínico (Admin)" : "Clinical (Admin)"}</span>
            </button>
            <button
              type="button"
              onClick={() => {
                changePortalMode("patient");
              }}
              className={`px-3 py-2 rounded-xl transition-all duration-300 flex items-center space-x-1.5 ${
                portalMode === "patient"
                  ? "bg-gradient-to-r from-pink-400 to-rose-500 text-white shadow-sm shadow-pink-500/20"
                  : "text-rose-800 dark:text-pink-300 hover:text-rose-950 dark:hover:text-white"
              }`}
            >
              <Heart size={12} className="fill-current text-rose-500 dark:text-pink-300" />
              <span>{lang === "es" ? "Paciente (Cliente)" : "Patient (Client)"}</span>
            </button>
          </div>

          {/* Dynamic Navigation Links based on Portal Mode */}
          {portalMode === "admin" ? (
            currentAdmin && (() => {
              const adminTabs = [
                { id: "agenda", label: t.agenda, icon: Calendar },
                { id: "patients", label: t.patients, icon: Users },
                { id: "odontogram", label: t.odontogram, icon: Activity },
                { id: "payments", label: t.payments, icon: DollarSign },
                { id: "availability", label: lang === "es" ? "Disponibilidad" : "Availability", icon: Clock },
                { id: "ai", label: t.aiAssistant, icon: Sparkles },
                { id: "sync", label: t.syncCenter, icon: Database },
                ...(currentAdmin?.role === "master" ? [{ id: "master-config", label: lang === "es" ? "Consola Master" : "Master Console", icon: Settings }] : [])
              ];
              const currentTab = adminTabs.find(tab => tab.id === activeTab) || adminTabs[0];
              const CurrentIcon = currentTab.icon;

              return (
                <div className="relative">
                  <button
                    type="button"
                    onClick={() => setShowAdminNavDropdown(!showAdminNavDropdown)}
                    className={`px-5 py-3 rounded-2xl border ${borderCol} ${
                      isDark ? "bg-pink-950/20 text-pink-300" : "bg-white text-rose-800"
                    } flex items-center justify-between min-w-[280px] text-xs font-bold uppercase tracking-wider shadow-sm hover:bg-pink-500/5 transition-all duration-300`}
                  >
                    <div className="flex items-center space-x-2.5">
                      <div className={`p-1.5 rounded-lg bg-gradient-to-r ${activeThemeObj.gradient} text-white`}>
                        <CurrentIcon size={14} />
                      </div>
                      <span>{currentTab.label}</span>
                    </div>
                    <span className="text-[10px] opacity-60 ml-2">▼</span>
                  </button>
                  
                  <AnimatePresence>
                    {showAdminNavDropdown && (
                      <>
                        {/* Backdrop to close the dropdown on click outside */}
                        <div 
                          className="fixed inset-0 z-40" 
                          onClick={() => setShowAdminNavDropdown(false)} 
                        />
                        <motion.div
                          initial={{ opacity: 0, y: 10, scale: 0.95 }}
                          animate={{ opacity: 1, y: 0, scale: 1 }}
                          exit={{ opacity: 0, y: 10, scale: 0.95 }}
                          transition={{ duration: 0.15 }}
                          className={`absolute left-1/2 -translate-x-1/2 md:left-0 md:translate-x-0 mt-2 w-full min-w-[280px] rounded-3xl border border-rose-100/50 dark:border-pink-500/10 bg-white dark:bg-[#201517] shadow-xl p-2 z-50 overflow-hidden max-h-[400px] overflow-y-auto`}
                        >
                          {adminTabs.map(tab => {
                            const Icon = tab.icon;
                            const isActive = activeTab === tab.id;
                            return (
                              <button
                                key={tab.id}
                                type="button"
                                onClick={() => {
                                  setActiveTab(tab.id as any);
                                  setSelectedTooth(null);
                                  setShowAdminNavDropdown(false);
                                }}
                                className={`w-full flex items-center space-x-3 px-4 py-3 text-xs font-bold uppercase tracking-wider rounded-2xl transition-all ${
                                  isActive 
                                    ? `bg-gradient-to-r ${activeThemeObj.gradient} text-white shadow-md ${activeThemeObj.shadow}` 
                                    : `${isDark ? `${activeThemeObj.accentTextLight} hover:bg-pink-950/40 hover:text-white` : "text-rose-800 hover:bg-rose-50 hover:text-rose-950"}`
                                }`}
                              >
                                <Icon size={14} className={isActive ? "text-white" : "text-pink-500"} />
                                <span>{tab.label}</span>
                              </button>
                            );
                          })}
                        </motion.div>
                      </>
                    )}
                  </AnimatePresence>
                </div>
              );
            })()
          ) : (
            loggedPatientId && (
              <nav className="flex flex-wrap items-center justify-center gap-1 bg-rose-100/30 dark:bg-pink-900/10 p-1.5 rounded-2xl border border-pink-500/5">
                {[
                  { id: "my-appointments", label: lang === "es" ? "Mis Citas" : "My Bookings", icon: Calendar },
                  { id: "my-odontogram", label: lang === "es" ? "Salud Dental" : "Dental Map", icon: Activity },
                  ...(activeFeatures.payments !== false ? [{ id: "my-billing", label: lang === "es" ? "Facturas y Pagos" : "Billing History", icon: DollarSign }] : []),
                  ...(activeFeatures.aiDiagnostic !== false ? [{ id: "my-ai", label: lang === "es" ? "Asistente IA" : "Aesthetic AI", icon: Sparkles }] : [])
                ].map(tab => {
                  const Icon = tab.icon;
                  const isActive = patientActiveTab === tab.id;
                  return (
                    <button
                      key={tab.id}
                      id={`nav-tab-patient-${tab.id}`}
                      onClick={() => {
                        setPatientActiveTab(tab.id as any);
                        setPatientSelectedTooth(null);
                      }}
                      className={`flex items-center space-x-2 px-4 py-2 rounded-xl text-xs font-bold transition-all uppercase tracking-wider ${
                        isActive 
                          ? `bg-gradient-to-r ${activeThemeObj.gradient} text-white shadow-md ${activeThemeObj.shadow}` 
                          : `${isDark ? `${activeThemeObj.accentTextLight} hover:bg-pink-950/40 hover:text-white` : "text-rose-800 hover:bg-rose-100/60 hover:text-rose-950"}`
                      }`}
                    >
                      <Icon size={14} />
                      <span>{tab.label}</span>
                    </button>
                  );
                })}
              </nav>
            )
          )}

          {/* Tool actions (Lang, Theme) */}
          <div className="flex items-center space-x-2">
            
            {/* Admin session state badge and logout */}
            {portalMode === "admin" && currentAdmin && (
              <div className={`flex items-center space-x-2 px-3 py-1.5 rounded-xl border ${borderCol} ${isDark ? "bg-pink-950/30" : "bg-rose-50/50"}`}>
                <div className="flex flex-col items-end text-[10px] hidden sm:flex">
                  <span className={`font-bold ${textTitle} max-w-[120px] truncate`}>{currentAdmin.name}</span>
                  <span className="text-pink-500 font-mono font-bold uppercase text-[8px]">{currentAdmin.role === "master" ? "Fundador (Master)" : "Admin"}</span>
                </div>
                <button
                  onClick={handleAdminLogout}
                  className="p-2 rounded-lg bg-rose-500/10 hover:bg-rose-500 hover:text-white text-rose-600 dark:text-rose-400 transition-all flex items-center space-x-1"
                  title="Cerrar Sesión Administrador"
                >
                  <LogOut size={12} />
                  <span className="text-[10px] font-bold uppercase sm:hidden">Salir</span>
                </button>
              </div>
            )}

            {/* Language Switcher */}
            <div className="relative">
              <button
                type="button"
                onClick={() => setShowLangDropdown(!showLangDropdown)}
                className={`p-2.5 rounded-xl border ${borderCol} ${isDark ? "bg-pink-950/20 text-pink-300" : "bg-white text-rose-800"} flex items-center space-x-1.5 text-xs font-bold shadow-sm hover:bg-pink-500/5 transition-all`}
              >
                <Globe size={14} className="text-pink-500" />
                <span className="uppercase">{lang}</span>
                <span className="text-[9px] opacity-60">▼</span>
              </button>
              
              <AnimatePresence>
                {showLangDropdown && (
                  <>
                    {/* Backdrop to close the dropdown on click outside */}
                    <div 
                      className="fixed inset-0 z-40" 
                      onClick={() => setShowLangDropdown(false)} 
                    />
                    <motion.div
                      initial={{ opacity: 0, y: 10, scale: 0.95 }}
                      animate={{ opacity: 1, y: 0, scale: 1 }}
                      exit={{ opacity: 0, y: 10, scale: 0.95 }}
                      transition={{ duration: 0.15 }}
                      className="absolute right-0 top-full mt-2 w-32 rounded-2xl border border-rose-100/50 bg-white dark:bg-[#201517] shadow-xl p-1.5 z-50 overflow-hidden"
                    >
                      {(["es", "en", "fr", "pt"] as const).map(l => (
                        <button
                          key={l}
                          type="button"
                          onClick={() => {
                            changeLanguage(l);
                            setShowLangDropdown(false);
                          }}
                          className={`w-full text-left px-3 py-1.5 text-xs font-bold uppercase rounded-lg transition-colors ${
                            lang === l 
                              ? "bg-rose-500 text-white" 
                              : "text-rose-900 dark:text-pink-100 hover:bg-rose-50 dark:hover:bg-pink-950/20"
                          }`}
                        >
                          {l === "es" ? "Español" : l === "en" ? "English" : l === "fr" ? "Français" : "Português"}
                        </button>
                      ))}
                    </motion.div>
                  </>
                )}
              </AnimatePresence>
            </div>

            {/* Light / Dark Mode Toggle */}
            <button
              id="theme-switcher"
              onClick={toggleTheme}
              className={`p-2.5 rounded-xl border ${borderCol} ${isDark ? "bg-pink-950/20 text-pink-300" : "bg-white text-rose-800"} shadow-sm transition-all hover:scale-105`}
              title={isDark ? "Modo Claro" : "Modo Oscuro"}
            >
              {isDark ? <Sun size={15} className="text-amber-400" /> : <Moon size={15} className="text-indigo-600" />}
            </button>
          </div>

        </div>
      </header>

      <main className="max-w-7xl mx-auto px-6 mt-8">
        {portalMode === "admin" && (
          !currentAdmin ? (
            <motion.div 
              initial={{ opacity: 0, y: 15 }} 
              animate={{ opacity: 1, y: 0 }}
              className="max-w-md mx-auto my-12"
            >
              <div className={bgCard}>
                <div className="flex flex-col items-center text-center space-y-3 pb-6 border-b border-pink-500/10 mb-6">
                  <div className="w-16 h-16 rounded-full bg-gradient-to-tr from-pink-400 to-rose-500 flex items-center justify-center text-white shadow-lg shadow-pink-500/20">
                    <Lock size={26} />
                  </div>
                  <div className="space-y-1">
                    <h2 className={`text-xl font-bold tracking-tight ${textTitle}`}>
                      {lang === "es" ? "Acceso Administrativo" : "Administrative Access"}
                    </h2>
                    <p className={`text-xs ${textSub}`}>
                      {lang === "es" ? "DianaSRL Odontología Stetic — Panel Clínico" : "DianaSRL Aesthetic Dentistry — Clinical Panel"}
                    </p>
                  </div>
                </div>

                <form onSubmit={handleAdminLogin} className="space-y-4">
                  <div>
                    <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5">
                      {lang === "es" ? "Usuario de Administrador" : "Administrator Username"}
                    </label>
                    <input
                      type="text"
                      required
                      placeholder={lang === "es" ? "Ej. Dianasrl@#" : "e.g. Dianasrl@#"}
                      value={adminUsernameInput}
                      onChange={e => setAdminUsernameInput(e.target.value)}
                      className={inputStyle}
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5">
                      {lang === "es" ? "Contraseña" : "Password"}
                    </label>
                    <input
                      type="password"
                      required
                      placeholder="••••••••"
                      value={adminPasswordInput}
                      onChange={e => setAdminPasswordInput(e.target.value)}
                      className={inputStyle}
                    />
                  </div>

                  <button
                    type="submit"
                    className="w-full py-3 bg-gradient-to-r from-pink-400 to-rose-500 hover:from-pink-500 hover:to-rose-600 text-white font-bold rounded-xl text-xs uppercase tracking-wider transition-all flex items-center justify-center space-x-2 shadow-md shadow-pink-500/10 hover:scale-[1.02]"
                  >
                    <Shield size={14} />
                    <span>{lang === "es" ? "Iniciar Sesión" : "Log In"}</span>
                  </button>
                </form>

                {/* Master Account Indicator for evaluation */}
                <div className={`mt-6 p-4 rounded-2xl border ${borderCol} ${isDark ? "bg-pink-950/10 text-pink-300" : "bg-rose-50/40 text-rose-950"} space-y-2`}>
                  <div className="flex items-center space-x-1.5 text-[10px] font-bold uppercase text-pink-500">
                    <Sparkles size={12} className="animate-pulse" />
                    <span>{lang === "es" ? "Autenticación Fundadora Master" : "Master Founder Auth"}</span>
                  </div>
                  <p className="text-[11px] leading-relaxed text-gray-500 dark:text-pink-300/70">
                    {lang === "es" 
                      ? "La única forma de registrar nuevas cuentas de administrador clínicas es iniciando sesión con la cuenta Master del Fundador de DianaSRL."
                      : "The only way to register new clinical admin accounts is by logging in with the DianaSRL Founder's Master account."}
                  </p>
                </div>
              </div>
            </motion.div>
          ) : (
            <>
            
            {/* Clinic Hero Brand Pitch */}
        <div className={`mb-8 p-6 rounded-3xl border ${borderCol} ${isDark ? "bg-gradient-to-r from-pink-950/10 via-[#2a1b1d] to-[#1c1214]" : "bg-gradient-to-r from-rose-50/40 via-[#fff8f9] to-[#fffbfc]"} flex flex-col md:flex-row items-center justify-between gap-6 relative overflow-hidden`}>
          <div className="absolute top-0 right-0 w-32 h-32 bg-pink-300/10 dark:bg-pink-500/5 rounded-full blur-2xl pointer-events-none" />
          <div className="space-y-1 z-10">
            <div className="inline-flex items-center space-x-2 px-3 py-1 rounded-full bg-pink-100 dark:bg-pink-950/30 text-pink-600 dark:text-pink-400 border border-pink-200/50 dark:border-pink-500/10">
              <span className="w-1.5 h-1.5 bg-emerald-500 rounded-full animate-ping" />
              <span className="text-[10px] font-bold uppercase tracking-widest font-mono">{t.clinicName} Premium System</span>
            </div>
            <h2 className={`text-2xl md:text-3xl font-extrabold tracking-tight ${textTitle}`}>{t.clinicTagline}</h2>
            <p className={`text-xs ${textSub} max-w-2xl leading-relaxed`}>
              {lang === "es" 
                ? "Diseño dental estético y clínico automatizado. Gestione fichas de pacientes, controle el odontograma interactivo, procese pagos de caja y sincronice toda su información con Google Sheets."
                : lang === "en"
                ? "Aesthetic and clinical dental automation. Manage patient records, control the interactive odontogram, process cash payments, and sync all data in real time to Google Sheets."
                : lang === "fr"
                ? "Automatisation dentaire clinique et esthétique. Gérez les dossiers, contrôlez l'odontogramme interactif, traitez les paiements et synchronisez le tout avec Google Sheets."
                : "Automação odontológica clínica e estética. Gerencie prontuários, controle o odontograma interativo, processe pagamentos de caixa e sincronize tudo com o Google Sheets."}
            </p>
          </div>

          <div className="flex flex-wrap gap-3 items-center z-10 shrink-0">
            <div className={`px-4 py-2 rounded-2xl border ${borderCol} ${isDark ? "bg-pink-950/20" : "bg-white"} text-center`}>
              <div className="text-xl font-bold text-pink-500">{patients.length}</div>
              <div className="text-[10px] font-bold uppercase tracking-wider text-gray-500">{lang === "es" ? "Pacientes" : "Patients"}</div>
            </div>
            <div className={`px-4 py-2 rounded-2xl border ${borderCol} ${isDark ? "bg-pink-950/20" : "bg-white"} text-center`}>
              <div className="text-xl font-bold text-emerald-500">{scheduledCount}</div>
              <div className="text-[10px] font-bold uppercase tracking-wider text-gray-500">{lang === "es" ? "Citas" : "Appointments"}</div>
            </div>
            <div className={`px-4 py-2 rounded-2xl border ${borderCol} ${isDark ? "bg-pink-950/20" : "bg-white"} text-center`}>
              <div className="text-xl font-bold text-pink-500">${totalCollectedAmt}</div>
              <div className="text-[10px] font-bold uppercase tracking-wider text-gray-500">{lang === "es" ? "Recaudado" : "Collected"}</div>
            </div>
          </div>
        </div>

        {/* Dynamic Screen Views */}
        <AnimatePresence mode="wait">
          
          {/* TAB 1: AGENDA & CITAS */}
          {activeTab === "agenda" && (
            <motion.div
              key="agenda-view"
              initial={{ opacity: 0, y: 15 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -15 }}
              className="grid grid-cols-1 lg:grid-cols-3 gap-8"
            >
              {/* Left Form: Booking */}
              <div className={`${bgCard} lg:col-span-1 h-fit`}>
                <div className="flex items-center space-x-2 border-b pb-4 mb-5 border-pink-500/10">
                  <Calendar className="text-pink-500" size={20} />
                  <h3 className={`text-lg font-bold ${textTitle}`}>{t.scheduleAppointment}</h3>
                </div>

                <form onSubmit={handleScheduleAppointment} className="space-y-4">
                  <div>
                    <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5">{t.selectPatient} *</label>
                    <select
                      value={newAppointment.patientId}
                      onChange={e => setNewAppointment(prev => ({ ...prev, patientId: e.target.value }))}
                      className={inputStyle}
                    >
                      <option value="">-- {lang === "es" ? "Seleccione un paciente" : "Select patient"} --</option>
                      {patients.map(p => (
                        <option key={p.id} value={p.id}>{p.name} (DNI: {p.dni})</option>
                      ))}
                    </select>
                  </div>

                  <div>
                    <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5">{t.selectDentist} *</label>
                    <select
                      value={newAppointment.dentistName}
                      onChange={e => setNewAppointment(prev => ({ ...prev, dentistName: e.target.value }))}
                      className={inputStyle}
                    >
                      <option value="Dra. Diana Rojas">Dra. Diana Rojas (Estética Dental)</option>
                      <option value="Dr. Marcos Soler">Dr. Marcos Soler (Cirugía & Implantes)</option>
                      <option value="Dra. Lucía Santos">Dra. Lucía Santos (Endodoncia)</option>
                    </select>
                  </div>

                  <div className="grid grid-cols-2 gap-4">
                    <div>
                      <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5">{t.appointmentDate} *</label>
                      <input
                        type="date"
                        value={newAppointment.date}
                        onChange={e => setNewAppointment(prev => ({ ...prev, date: e.target.value }))}
                        className={inputStyle}
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5">{t.appointmentTime} *</label>
                      <input
                        type="time"
                        value={newAppointment.time}
                        onChange={e => setNewAppointment(prev => ({ ...prev, time: e.target.value }))}
                        className={inputStyle}
                      />
                    </div>
                  </div>

                  <div>
                    <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5">{t.treatmentType} *</label>
                    <select
                      value={newAppointment.treatmentType}
                      onChange={e => setNewAppointment(prev => ({ ...prev, treatmentType: e.target.value }))}
                      className={inputStyle}
                    >
                      <option value="">-- {lang === "es" ? "Seleccione tratamiento" : "Select treatment"} --</option>
                      <option value="Resina Estética">Resina Estética / Obturación</option>
                      <option value="Carillas de Disilicato">Carillas de Disilicato de Litio</option>
                      <option value="Blanqueamiento Láser">Blanqueamiento Láser Dental</option>
                      <option value="Endodoncia">Endodoncia / Tratamiento de Conducto</option>
                      <option value="Limpieza Profiláctica">Limpieza Profiláctica & Ultrasonido</option>
                      <option value="Implante Titanio">Implante Dental de Titanio</option>
                      <option value="Corona Zirconio">Corona de Zirconio Estética</option>
                    </select>
                  </div>

                  <div>
                    <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5">{t.appointmentNotes}</label>
                    <textarea
                      rows={3}
                      placeholder={lang === "es" ? "Indique el motivo específico o notas clínicas..." : "Enter reasons or clinical notes..."}
                      value={newAppointment.notes}
                      onChange={e => setNewAppointment(prev => ({ ...prev, notes: e.target.value }))}
                      className={`${inputStyle} resize-none`}
                    />
                  </div>

                  <button
                    type="submit"
                    className="w-full py-3 bg-gradient-to-r from-pink-400 to-rose-500 hover:from-pink-500 hover:to-rose-600 text-white font-bold rounded-xl text-xs uppercase tracking-widest shadow-md shadow-pink-500/10 transition-all flex items-center justify-center space-x-2"
                  >
                    <PlusCircle size={15} />
                    <span>{t.bookNow}</span>
                  </button>
                </form>
              </div>

              {/* Right: List of Bookings */}
              <div className={`${bgCard} lg:col-span-2`}>
                <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b pb-4 mb-5 border-pink-500/10">
                  <div className="flex items-center space-x-2">
                    <Clock className="text-pink-500" size={20} />
                    <h3 className={`text-lg font-bold ${textTitle}`}>{t.upcomingAppointments}</h3>
                  </div>
                  <span className="text-xs bg-pink-100 dark:bg-pink-950/40 text-pink-600 dark:text-pink-400 font-bold px-3 py-1 rounded-full border border-pink-200/40 dark:border-pink-500/10">
                    {scheduledCount} {lang === "es" ? "Citas Activas" : "Active Bookings"}
                  </span>
                </div>

                {appointments.length === 0 ? (
                  <div className="text-center py-16 text-gray-400 space-y-3">
                    <Calendar className="mx-auto text-gray-300" size={48} />
                    <p className="text-sm font-medium">{t.noAppointments}</p>
                  </div>
                ) : (
                  <div className="space-y-4 max-h-[580px] overflow-y-auto pr-2 custom-scrollbar">
                    {appointments.map(app => {
                      const statusColors = {
                        Scheduled: "bg-amber-100 text-amber-700 dark:bg-amber-950/30 dark:text-amber-400 border-amber-200/50 dark:border-amber-500/10",
                        Completed: "bg-emerald-100 text-emerald-700 dark:bg-emerald-950/30 dark:text-emerald-400 border-emerald-200/50 dark:border-emerald-500/10",
                        Cancelled: "bg-red-100 text-red-700 dark:bg-red-950/30 dark:text-red-400 border-red-200/50 dark:border-red-500/10"
                      };

                      return (
                        <div 
                          key={app.id} 
                          className={`p-5 rounded-2xl border ${borderCol} ${isDark ? "bg-[#1f1618] hover:bg-[#251a1d]" : "bg-rose-50/10 hover:bg-rose-50/40"} transition-all flex flex-col md:flex-row justify-between gap-4`}
                        >
                          <div className="space-y-2">
                            <div className="flex items-center space-x-3">
                              <span className={`px-2.5 py-0.5 rounded-full text-[10px] font-extrabold uppercase border ${statusColors[app.status]}`}>
                                {app.status === "Scheduled" ? t.statusScheduled : app.status === "Completed" ? t.statusCompleted : t.statusCancelled}
                              </span>
                              <span className="text-[10px] font-bold text-pink-500 uppercase tracking-widest font-mono">{app.dentistName}</span>
                            </div>

                            <h4 className={`text-base font-bold ${textTitle}`}>{app.patientName}</h4>
                            <p className="text-xs text-gray-500 dark:text-pink-300/60 font-medium">
                              {lang === "es" ? "Procedimiento:" : "Procedure:"} <strong className="text-pink-500">{app.treatmentType}</strong>
                            </p>
                            
                            {app.notes && (
                              <p className="text-xs bg-white dark:bg-[#1a1113] p-3 rounded-xl border border-pink-500/5 text-gray-500 italic mt-1 leading-relaxed">
                                "{app.notes}"
                              </p>
                            )}
                          </div>

                          <div className="flex flex-row md:flex-col justify-between items-end shrink-0 gap-4">
                            <div className="text-right">
                              <div className={`text-sm font-bold ${textTitle} flex items-center justify-end space-x-1`}>
                                <Calendar size={13} className="text-pink-400" />
                                <span>{app.date}</span>
                              </div>
                              <div className="text-xs text-gray-400 dark:text-pink-300/40 font-bold flex items-center justify-end space-x-1 mt-0.5">
                                <Clock size={13} className="text-pink-400" />
                                <span>{app.time} HS</span>
                              </div>
                            </div>

                            {app.status === "Scheduled" && (
                              <div className="flex items-center space-x-2">
                                <button
                                  onClick={() => handleUpdateAppointmentStatus(app.id, "Completed")}
                                  className="px-3 py-1.5 bg-emerald-500 hover:bg-emerald-600 text-white font-bold text-[10px] rounded-lg shadow-sm transition-colors uppercase tracking-wider"
                                >
                                  {lang === "es" ? "Completar" : "Complete"}
                                </button>
                                <button
                                  onClick={() => handleUpdateAppointmentStatus(app.id, "Cancelled")}
                                  className="px-3 py-1.5 bg-red-100 hover:bg-red-200 text-red-600 font-bold text-[10px] rounded-lg transition-colors uppercase tracking-wider dark:bg-red-950/20 dark:text-red-400"
                                >
                                  {lang === "es" ? "Cancelar" : "Cancel"}
                                </button>
                              </div>
                            )}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
            </motion.div>
          )}

          {/* TAB 2: PACIENTES */}
          {activeTab === "patients" && (
            <motion.div
              key="patients-view"
              initial={{ opacity: 0, y: 15 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -15 }}
              className="grid grid-cols-1 lg:grid-cols-3 gap-8"
            >
              {/* Left Column: Register patient */}
              <div className={`${bgCard} lg:col-span-1 h-fit`}>
                <div className="flex items-center space-x-2 border-b pb-4 mb-5 border-pink-500/10">
                  <Users className="text-pink-500" size={20} />
                  <h3 className={`text-lg font-bold ${textTitle}`}>{t.addPatient}</h3>
                </div>

                <form onSubmit={handleRegisterPatient} className="space-y-4">
                  <div>
                    <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5">{t.patientName} *</label>
                    <input
                      type="text"
                      placeholder="e.g. Diana María Pérez"
                      value={newPatient.name}
                      onChange={e => setNewPatient(prev => ({ ...prev, name: e.target.value }))}
                      className={inputStyle}
                      required
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5">{t.patientDni} *</label>
                    <input
                      type="text"
                      placeholder="e.g. 001-1234567-8"
                      value={newPatient.dni}
                      onChange={e => setNewPatient(prev => ({ ...prev, dni: e.target.value }))}
                      className={inputStyle}
                      required
                    />
                  </div>

                  <div className="grid grid-cols-2 gap-4">
                    <div>
                      <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5">{t.patientPhone}</label>
                      <input
                        type="text"
                        placeholder="e.g. 809-555-0123"
                        value={newPatient.phone}
                        onChange={e => setNewPatient(prev => ({ ...prev, phone: e.target.value }))}
                        className={inputStyle}
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5">{t.patientDob}</label>
                      <input
                        type="date"
                        value={newPatient.dob}
                        onChange={e => setNewPatient(prev => ({ ...prev, dob: e.target.value }))}
                        className={inputStyle}
                      />
                    </div>
                  </div>

                  <div>
                    <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5">{t.patientEmail}</label>
                    <input
                      type="email"
                      placeholder="e.g. diana@perez.com"
                      value={newPatient.email}
                      onChange={e => setNewPatient(prev => ({ ...prev, email: e.target.value }))}
                      className={inputStyle}
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5">{t.patientNotes}</label>
                    <textarea
                      rows={4}
                      placeholder={lang === "es" ? "Alergias, medicamentos continuos, antecedentes..." : "Allergies, ongoing treatments, notes..."}
                      value={newPatient.notes}
                      onChange={e => setNewPatient(prev => ({ ...prev, notes: e.target.value }))}
                      className={`${inputStyle} resize-none`}
                    />
                  </div>

                  <button
                    type="submit"
                    className="w-full py-3 bg-gradient-to-r from-pink-400 to-rose-500 hover:from-pink-500 hover:to-rose-600 text-white font-bold rounded-xl text-xs uppercase tracking-widest shadow-md shadow-pink-500/10 transition-all flex items-center justify-center space-x-2"
                  >
                    <PlusCircle size={15} />
                    <span>{t.registerPatient}</span>
                  </button>
                </form>
              </div>

              {/* Right Column: Listing and Search */}
              <div className={`${bgCard} lg:col-span-2`}>
                <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b pb-4 mb-5 border-pink-500/10">
                  <div className="flex items-center space-x-2">
                    <Users className="text-pink-500" size={20} />
                    <h3 className={`text-lg font-bold ${textTitle}`}>{t.patientList}</h3>
                  </div>
                  <span className="text-xs bg-rose-100 dark:bg-pink-950/40 text-pink-600 dark:text-pink-400 font-bold px-3 py-1 rounded-full border border-pink-200/40 dark:border-pink-500/10">
                    {patients.length} {lang === "es" ? "Registrados" : "Registered"}
                  </span>
                </div>

                {/* Patient cards display */}
                {patients.length === 0 ? (
                  <div className="text-center py-16 text-gray-400 space-y-3">
                    <Users className="mx-auto text-gray-300" size={48} />
                    <p className="text-sm font-medium">{t.noPatients}</p>
                  </div>
                ) : (
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4 max-h-[580px] overflow-y-auto pr-2 custom-scrollbar">
                    {patients.map(p => {
                      const hasToothData = (toothConditions[p.id] || []).length > 0;
                      return (
                        <div 
                          key={p.id} 
                          onClick={() => {
                            setSelectedPatientId(p.id);
                            setActiveTab("odontogram");
                          }}
                          className={`p-5 rounded-2xl border cursor-pointer transition-all ${
                            selectedPatientId === p.id 
                              ? "bg-pink-500/10 border-pink-400" 
                              : `hover:bg-rose-50/40 border-rose-100/50 dark:border-pink-500/5 ${isDark ? "bg-[#1f1618] hover:bg-[#251a1d]" : "bg-rose-50/10"}`
                          }`}
                        >
                          <div className="flex items-start justify-between">
                            <div className="space-y-1">
                              <h4 className={`text-base font-bold ${textTitle}`}>{p.name}</h4>
                              <p className="text-xs text-gray-400 font-mono">DNI: {p.dni}</p>
                            </div>
                            <div className="w-8 h-8 rounded-full bg-pink-100 dark:bg-pink-950/40 flex items-center justify-center text-pink-500 text-xs font-bold shrink-0">
                              {p.name.charAt(0)}
                            </div>
                          </div>

                          <div className="mt-4 grid grid-cols-2 gap-3 text-xs font-semibold text-gray-500 dark:text-pink-300/60 border-t pt-3 border-pink-500/5">
                            <div>
                              <span className="text-[10px] text-gray-400 block uppercase font-bold">{lang === "es" ? "Teléfono" : "Phone"}</span>
                              <span className={isDark ? "text-pink-100" : "text-rose-950"}>
                                {p.phone || "--"}{" "}
                                {p.preferredContact && (
                                  <span className="text-[9px] bg-pink-500/10 text-pink-500 px-1 py-0.5 rounded font-bold uppercase ml-1">
                                    {p.preferredContact}
                                  </span>
                                )}
                              </span>
                            </div>
                            <div>
                              <span className="text-[10px] text-gray-400 block uppercase font-bold">{lang === "es" ? "F. Nacimiento" : "Birth Date"}</span>
                              <span className={isDark ? "text-pink-100" : "text-rose-950"}>{p.dob || "--"}</span>
                            </div>
                            
                            {p.dentalInsurance && (
                              <div className="col-span-1">
                                <span className="text-[10px] text-gray-400 block uppercase font-bold">{lang === "es" ? "Seguro Dental" : "Dental Insurance"}</span>
                                <span className="text-pink-500 text-xs font-extrabold">{p.dentalInsurance}</span>
                              </div>
                            )}

                            {p.consultationReason && (
                              <div className={p.dentalInsurance ? "col-span-1" : "col-span-2"}>
                                <span className="text-[10px] text-gray-400 block uppercase font-bold">{lang === "es" ? "Motivo Consulta" : "Consultation Reason"}</span>
                                <span className={isDark ? "text-pink-300" : "text-rose-900"}>{p.consultationReason}</span>
                              </div>
                            )}
                          </div>

                          {/* Medical Alerts (Allergies or Conditions) */}
                          {(p.allergies || p.medicalConditions) && (
                            <div className="mt-3 p-2.5 rounded-xl bg-amber-500/5 border border-amber-500/10 space-y-1">
                              {p.allergies && (
                                <div className="text-[11px] leading-tight text-amber-700 dark:text-amber-400">
                                  <strong className="font-bold uppercase tracking-wider text-[9px]">{lang === "es" ? "Alergias:" : "Allergies:"} </strong> 
                                  {p.allergies}
                                </div>
                              )}
                              {p.medicalConditions && (
                                <div className="text-[11px] leading-tight text-amber-700 dark:text-amber-400">
                                  <strong className="font-bold uppercase tracking-wider text-[9px]">{lang === "es" ? "Condiciones:" : "Conditions:"} </strong> 
                                  {p.medicalConditions}
                                </div>
                              )}
                            </div>
                          )}

                          {p.notes && (
                            <p className="mt-3 text-[11px] text-rose-700/80 dark:text-pink-300/60 bg-pink-50/30 dark:bg-[#1a1113] p-2 rounded-lg border border-pink-500/5 italic leading-relaxed line-clamp-2">
                              "{p.notes}"
                            </p>
                          )}

                          {p.address && (
                            <div className="mt-2 text-[10px] text-gray-500 dark:text-pink-300/40 flex items-center space-x-1">
                              <span className="font-extrabold uppercase">{lang === "es" ? "📍 Dirección:" : "📍 Address:"}</span>
                              <span className="truncate">{p.address}</span>
                            </div>
                          )}

                          <div className="mt-3 flex items-center justify-between text-[10px] font-bold">
                            <span className={`inline-flex items-center space-x-1 ${hasToothData ? "text-emerald-500" : "text-gray-400"}`}>
                              <CheckCircle size={10} />
                              <span>{hasToothData ? "Ficha dental activa" : "Sin odontograma"}</span>
                            </span>
                            <span className="text-pink-500 uppercase tracking-wider flex items-center hover:underline">
                              <span>Odontograma</span>
                              <ChevronRight size={10} className="ml-0.5" />
                            </span>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
            </motion.div>
          )}

          {/* TAB 3: ODONTOGRAMA INTERACTIVO */}
          {activeTab === "odontogram" && (
            <motion.div
              key="odontogram-view"
              initial={{ opacity: 0, y: 15 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -15 }}
              className="space-y-8"
            >
              {/* Odontogram Header Selection */}
              <div className={`${bgCard} p-6 flex flex-col md:flex-row items-center justify-between gap-4`}>
                <div className="space-y-1">
                  <span className="text-[10px] font-extrabold uppercase text-pink-500 font-mono tracking-widest">{t.odontogram}</span>
                  <h3 className={`text-xl font-bold ${textTitle}`}>{t.chartTitle}</h3>
                  <p className={`text-xs ${textSub}`}>{t.chartSubtitle}</p>
                </div>

                <div className="flex items-center space-x-3 shrink-0">
                  <label className="text-xs font-bold uppercase text-gray-500">{lang === "es" ? "Paciente Activo:" : "Active Patient:"}</label>
                  <select
                    value={selectedPatientId}
                    onChange={e => setSelectedPatientId(e.target.value)}
                    className={`${inputStyle} w-64`}
                  >
                    {patients.map(p => (
                      <option key={p.id} value={p.id}>{p.name} (DNI: {p.dni})</option>
                    ))}
                  </select>
                </div>
              </div>

              {/* Teeth Canvas / Layout Panel */}
              <div className="grid grid-cols-1 lg:grid-cols-4 gap-8">
                
                {/* Visual Upper/Lower Teeth Layout */}
                <div className={`${bgCard} lg:col-span-3 space-y-8 overflow-x-auto custom-scrollbar`}>
                  
                  {/* Arcada Superior (Upper Arc) */}
                  <div className="space-y-3">
                    <div className="flex items-center justify-between border-b border-pink-500/5 pb-2">
                      <span className="text-xs font-bold text-pink-500 tracking-wider uppercase flex items-center space-x-1">
                        <Layers size={13} />
                        <span>{lang === "es" ? "Arcada Superior" : "Upper Teeth Arc"}</span>
                      </span>
                      <span className="text-[10px] text-gray-400 font-mono font-bold">PIEZAS 1 - 16</span>
                    </div>

                    <div className="grid grid-cols-16 gap-2 min-w-[700px] py-4 bg-rose-50/10 dark:bg-pink-950/10 rounded-2xl border border-pink-500/5 px-4 justify-items-center">
                      {Array.from({ length: 16 }, (_, i) => {
                        const toothNum = i + 1;
                        const patientConditions = toothConditions[selectedPatientId] || [];
                        const condition = patientConditions.find(c => c.toothNumber === toothNum);
                        const status = condition?.status || "Healthy";

                        // Map tooth statuses to color indicators
                        const statusTheme = {
                          Healthy: "bg-emerald-100 border-emerald-300 text-emerald-700 dark:bg-emerald-950/30 dark:text-emerald-400 dark:border-emerald-500/20",
                          Caries: "bg-red-100 border-red-300 text-red-700 dark:bg-red-950/40 dark:text-red-400 dark:border-red-500/30 animate-pulse",
                          Endo: "bg-purple-100 border-purple-300 text-purple-700 dark:bg-purple-950/30 dark:text-purple-400 dark:border-purple-500/20",
                          Missing: "bg-dashed bg-gray-100 border-dashed border-gray-300 text-gray-400 dark:bg-gray-950/30 dark:border-gray-800",
                          Crown: "bg-amber-100 border-amber-300 text-amber-700 dark:bg-amber-950/30 dark:text-amber-400 dark:border-amber-500/20",
                          Filling: "bg-blue-100 border-blue-300 text-blue-700 dark:bg-blue-950/30 dark:text-blue-400 dark:border-blue-500/20"
                        }[status];

                        const isActive = selectedTooth === toothNum;

                        return (
                          <div
                            key={toothNum}
                            onClick={() => setSelectedTooth(toothNum)}
                            className={`w-11 h-16 rounded-xl border flex flex-col items-center justify-between p-1.5 cursor-pointer transition-all ${statusTheme} ${
                              isActive ? "ring-4 ring-pink-500 ring-offset-2 dark:ring-offset-[#181113]" : "hover:scale-105"
                            }`}
                          >
                            <span className="text-[9px] font-bold font-mono text-gray-400">{toothNum}</span>
                            {/* Visual Tooth SVG Accent */}
                            <div className="w-5 h-6 flex items-center justify-center opacity-85 my-0.5">
                              <span className="text-base font-bold select-none">{status === "Missing" ? "✕" : "🦷"}</span>
                            </div>
                            <span className="text-[7px] font-extrabold uppercase tracking-tighter text-center line-clamp-1 block leading-tight">
                              {status === "Healthy" ? "Sano" : status === "Caries" ? "Caries" : status === "Endo" ? "Endo" : status === "Missing" ? "Ausente" : status === "Crown" ? "Corona" : "Resina"}
                            </span>
                          </div>
                        );
                      })}
                    </div>
                  </div>

                  {/* Arcada Inferior (Lower Arc) */}
                  <div className="space-y-3">
                    <div className="flex items-center justify-between border-b border-pink-500/5 pb-2">
                      <span className="text-xs font-bold text-pink-500 tracking-wider uppercase flex items-center space-x-1">
                        <Layers size={13} />
                        <span>{lang === "es" ? "Arcada Inferior" : "Lower Teeth Arc"}</span>
                      </span>
                      <span className="text-[10px] text-gray-400 font-mono font-bold">PIEZAS 17 - 32</span>
                    </div>

                    {/* Lower teeth list in anatomically matched order (32 to 17 left-right mirror) */}
                    <div className="grid grid-cols-16 gap-2 min-w-[700px] py-4 bg-rose-50/10 dark:bg-pink-950/10 rounded-2xl border border-pink-500/5 px-4 justify-items-center">
                      {Array.from({ length: 16 }, (_, i) => {
                        const toothNum = 32 - i;
                        const patientConditions = toothConditions[selectedPatientId] || [];
                        const condition = patientConditions.find(c => c.toothNumber === toothNum);
                        const status = condition?.status || "Healthy";

                        const statusTheme = {
                          Healthy: "bg-emerald-100 border-emerald-300 text-emerald-700 dark:bg-emerald-950/30 dark:text-emerald-400 dark:border-emerald-500/20",
                          Caries: "bg-red-100 border-red-300 text-red-700 dark:bg-red-950/40 dark:text-red-400 dark:border-red-500/30 animate-pulse",
                          Endo: "bg-purple-100 border-purple-300 text-purple-700 dark:bg-purple-950/30 dark:text-purple-400 dark:border-purple-500/20",
                          Missing: "bg-dashed bg-gray-100 border-dashed border-gray-300 text-gray-400 dark:bg-gray-950/30 dark:border-gray-800",
                          Crown: "bg-amber-100 border-amber-300 text-amber-700 dark:bg-amber-950/30 dark:text-amber-400 dark:border-amber-500/20",
                          Filling: "bg-blue-100 border-blue-300 text-blue-700 dark:bg-blue-950/30 dark:text-blue-400 dark:border-blue-500/20"
                        }[status];

                        const isActive = selectedTooth === toothNum;

                        return (
                          <div
                            key={toothNum}
                            onClick={() => setSelectedTooth(toothNum)}
                            className={`w-11 h-16 rounded-xl border flex flex-col items-center justify-between p-1.5 cursor-pointer transition-all ${statusTheme} ${
                              isActive ? "ring-4 ring-pink-500 ring-offset-2 dark:ring-offset-[#181113]" : "hover:scale-105"
                            }`}
                          >
                            <span className="text-[7px] font-extrabold uppercase tracking-tighter text-center line-clamp-1 block leading-tight">
                              {status === "Healthy" ? "Sano" : status === "Caries" ? "Caries" : status === "Endo" ? "Endo" : status === "Missing" ? "Ausente" : status === "Crown" ? "Corona" : "Resina"}
                            </span>
                            <div className="w-5 h-6 flex items-center justify-center opacity-85 my-0.5">
                              <span className="text-base font-bold select-none">{status === "Missing" ? "✕" : "🦷"}</span>
                            </div>
                            <span className="text-[9px] font-bold font-mono text-gray-400">{toothNum}</span>
                          </div>
                        );
                      })}
                    </div>
                  </div>

                  {/* Diagnostic Legend Code */}
                  <div className="pt-4 border-t border-pink-500/5 grid grid-cols-3 md:grid-cols-6 gap-3 justify-items-center">
                    <div className="flex items-center space-x-2 text-xs font-bold text-gray-500">
                      <span className="w-4 h-4 bg-emerald-500 rounded-lg flex items-center justify-center text-[10px] text-white">🦷</span>
                      <span>{t.stateHealthy}</span>
                    </div>
                    <div className="flex items-center space-x-2 text-xs font-bold text-gray-500">
                      <span className="w-4 h-4 bg-red-500 rounded-lg flex items-center justify-center text-[10px] text-white">🦷</span>
                      <span>{t.stateCaries}</span>
                    </div>
                    <div className="flex items-center space-x-2 text-xs font-bold text-gray-500">
                      <span className="w-4 h-4 bg-purple-500 rounded-lg flex items-center justify-center text-[10px] text-white">🦷</span>
                      <span>{t.stateEndo}</span>
                    </div>
                    <div className="flex items-center space-x-2 text-xs font-bold text-gray-500">
                      <span className="w-4 h-4 bg-gray-300 rounded-lg flex items-center justify-center text-[10px] text-white">✕</span>
                      <span>{t.stateMissing}</span>
                    </div>
                    <div className="flex items-center space-x-2 text-xs font-bold text-gray-500">
                      <span className="w-4 h-4 bg-amber-500 rounded-lg flex items-center justify-center text-[10px] text-white">👑</span>
                      <span>{t.stateCrown}</span>
                    </div>
                    <div className="flex items-center space-x-2 text-xs font-bold text-gray-500">
                      <span className="w-4 h-4 bg-blue-500 rounded-lg flex items-center justify-center text-[10px] text-white">🦷</span>
                      <span>{t.stateFilling}</span>
                    </div>
                  </div>

                </div>

                {/* Right Panel: Selected Tooth Treatment Config */}
                <div className={`${bgCard} lg:col-span-1 flex flex-col justify-between h-full min-h-[400px]`}>
                  {selectedTooth === null ? (
                    <div className="my-auto text-center py-10 space-y-4 text-gray-400">
                      <div className="w-16 h-16 bg-rose-500/10 text-rose-500 rounded-full flex items-center justify-center mx-auto text-2xl">🦷</div>
                      <p className="text-xs font-bold leading-relaxed">
                        {lang === "es" 
                          ? "Seleccione un diente de las arcadas para registrar tratamientos, restauraciones o caries." 
                          : "Select a tooth from the charts to record clinical diagnoses, treatments or restorations."}
                      </p>
                    </div>
                  ) : (
                    <div className="space-y-6">
                      <div className="border-b border-pink-500/5 pb-4">
                        <div className="text-[10px] font-bold text-pink-500 tracking-wider font-mono uppercase">{t.clinicalRecord}</div>
                        <h4 className={`text-lg font-bold ${textTitle} mt-0.5`}>{t.selectTooth} #{selectedTooth}</h4>
                      </div>

                      <div className="space-y-4">
                        <label className="block text-xs font-bold uppercase text-gray-500">{t.toothState}</label>
                        
                        <div className="grid grid-cols-1 gap-2.5">
                          {[
                            { id: "Healthy", label: t.stateHealthy, icon: "🟢" },
                            { id: "Caries", label: t.stateCaries, icon: "🔴" },
                            { id: "Endo", label: t.stateEndo, icon: "🟣" },
                            { id: "Missing", label: t.stateMissing, icon: "✕" },
                            { id: "Crown", label: t.stateCrown, icon: "👑" },
                            { id: "Filling", label: t.stateFilling, icon: "🔵" }
                          ].map(state => (
                            <button
                              key={state.id}
                              onClick={() => handleApplyToothCondition(state.id as any)}
                              className={`w-full px-4 py-3 border rounded-xl font-bold text-xs uppercase flex items-center space-x-3 transition-all hover:scale-[1.02] ${
                                isDark 
                                  ? "bg-[#1f1618] border-pink-500/10 hover:bg-[#281c1f]" 
                                  : "bg-rose-50/20 border-rose-100 hover:bg-rose-50/60"
                              }`}
                            >
                              <span className="text-sm select-none">{state.icon}</span>
                              <span className={isDark ? "text-pink-100" : "text-rose-950"}>{state.label}</span>
                            </button>
                          ))}
                        </div>
                      </div>

                      <div className="pt-4 border-t border-pink-500/5 text-center">
                        <button
                          onClick={() => setSelectedTooth(null)}
                          className="text-xs text-pink-500 hover:text-pink-600 font-bold transition-all underline"
                        >
                          {lang === "es" ? "Cancelar Selección" : "Cancel Selection"}
                        </button>
                      </div>
                    </div>
                  )}

                  {/* Active Patient Teeth summary list */}
                  <div className="mt-6 pt-4 border-t border-pink-500/5 space-y-3">
                    <h5 className="text-[10px] font-bold uppercase tracking-wider text-gray-400">
                      {lang === "es" ? "Tratamientos Registrados" : "Diagnosed Issues"}
                    </h5>
                    <div className="space-y-2 max-h-40 overflow-y-auto pr-1 custom-scrollbar">
                      {patients.find(p => p.id === selectedPatientId) && (toothConditions[selectedPatientId] || []).length === 0 && (
                        <span className="text-[10px] text-gray-400 italic block py-2">{lang === "es" ? "Dentadura sana sin anomalías registradas." : "Denture healthy without registered entries."}</span>
                      )}
                      {(toothConditions[selectedPatientId] || []).map(cond => (
                        <div key={cond.toothNumber} className="flex items-center justify-between text-[11px] bg-rose-500/5 px-2.5 py-1.5 rounded-lg border border-pink-500/5">
                          <span className="font-bold text-pink-500">Diente #{cond.toothNumber}</span>
                          <span className="uppercase font-extrabold text-[9px] tracking-wider text-gray-500">{cond.status}</span>
                        </div>
                      ))}
                    </div>
                  </div>

                </div>

              </div>
            </motion.div>
          )}

          {/* TAB 4: FACTURACIÓN Y CAJA */}
          {activeTab === "payments" && (
            <motion.div
              key="payments-view"
              initial={{ opacity: 0, y: 15 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -15 }}
              className="grid grid-cols-1 lg:grid-cols-3 gap-8"
            >
              {/* Header Navigation Tab Selector */}
              <div className="col-span-1 lg:col-span-3 flex flex-col md:flex-row items-center justify-between gap-4 border-b pb-4 border-pink-500/10">
                <div className="space-y-1">
                  <h3 className={`text-xl font-extrabold tracking-tight ${textTitle}`}>
                    {lang === "es" ? "Facturación, Cobros y Pasarelas de Pago" : "Billing, Invoices & Payment Gateways"}
                  </h3>
                  <p className={`text-xs ${textSub}`}>
                    {lang === "es" ? "Gestione los cobros diarios y configure las pasarelas o cuentas para recibir los fondos de los clientes." : "Manage daily invoices and configure payment gateways or deposit accounts for patient funds."}
                  </p>
                </div>
                <div className="flex bg-rose-100/40 dark:bg-pink-900/15 p-1 rounded-2xl border border-pink-500/10 select-none font-bold text-[10px] uppercase tracking-wider gap-0.5 shrink-0">
                  <button
                    type="button"
                    onClick={() => setPaymentsSubTab("caja")}
                    className={`px-4 py-2 rounded-xl transition-all duration-300 flex items-center space-x-1.5 ${
                      paymentsSubTab === "caja"
                        ? "bg-gradient-to-r from-pink-400 to-rose-500 text-white shadow-sm shadow-pink-500/20"
                        : "text-rose-800 dark:text-pink-300 hover:text-rose-950 dark:hover:text-white"
                    }`}
                  >
                    <TrendingUp size={12} />
                    <span>{lang === "es" ? "Caja Registradora" : "Cashier Register"}</span>
                  </button>
                  <button
                    type="button"
                    onClick={() => setPaymentsSubTab("config")}
                    className={`px-4 py-2 rounded-xl transition-all duration-300 flex items-center space-x-1.5 ${
                      paymentsSubTab === "config"
                        ? "bg-gradient-to-r from-pink-400 to-rose-500 text-white shadow-sm shadow-pink-500/20"
                        : "text-rose-800 dark:text-pink-300 hover:text-rose-950 dark:hover:text-white"
                    }`}
                  >
                    <CreditCard size={12} />
                    <span>{lang === "es" ? "Métodos de Pago" : "Payment Methods"}</span>
                  </button>
                </div>
              </div>

              {paymentsSubTab === "config" ? (
                <div className="col-span-1 lg:col-span-3 space-y-6">
                  {/* Permissions Banner */}
                  {!currentAdmin || currentAdmin.role !== "master" ? (
                    <div className="p-4 rounded-2xl bg-amber-500/10 border border-amber-500/20 flex items-start space-x-3">
                      <Lock className="text-amber-500 shrink-0 mt-0.5" size={18} />
                      <div>
                        <h4 className="text-xs font-extrabold text-amber-500 uppercase tracking-wider">Modo de Solo Lectura</h4>
                        <p className="text-xs text-gray-400 mt-1">
                          Solo la fundadora y administradora Master (<strong>Dra. Diana Rojas</strong>) tiene privilegios de edición para habilitar, deshabilitar o reconfigurar los enlaces de depósito y cuentas donde cae el dinero de los clientes.
                        </p>
                      </div>
                    </div>
                  ) : (
                    <div className="p-5 rounded-2xl bg-pink-500/10 border border-pink-500/20 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4">
                      <div className="flex items-start space-x-3">
                        <Shield className="text-pink-500 shrink-0 mt-0.5" size={20} />
                        <div>
                          <h4 className="text-xs font-extrabold text-pink-500 uppercase tracking-wider">Control de Pasarelas de Pago Autorizadas</h4>
                          <p className="text-xs text-gray-400 mt-0.5">
                            Usted tiene acceso total como Administradora Fundadora para cambiar cuentas, enlaces de PayPal/Stripe y habilitar opciones para sus clientes.
                          </p>
                        </div>
                      </div>
                      <button
                        type="button"
                        onClick={handleAddGateway}
                        className="px-4 py-2.5 bg-gradient-to-r from-pink-400 to-rose-500 hover:from-pink-500 hover:to-rose-600 text-white rounded-xl text-xs font-extrabold uppercase tracking-wider transition-all flex items-center justify-center space-x-2 shrink-0 shadow-md shadow-pink-500/10"
                      >
                        <Plus size={14} />
                        <span>{lang === "es" ? "Añadir Método" : "Add Method"}</span>
                      </button>
                    </div>
                  )}

                  {/* Gateways Grid */}
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                    {gateways.map((gw) => {
                      const isDefault = ["banco_popular", "banco_bhd", "paypal", "stripe_card"].includes(gw.id);
                      const canEdit = currentAdmin && currentAdmin.role === "master";

                      return (
                        <div
                          key={gw.id}
                          className={`p-6 rounded-3xl border ${borderCol} ${
                            gw.enabled 
                              ? (isDark ? "bg-[#1f1618] border-pink-500/25 shadow-lg shadow-pink-500/2" : "bg-white border-pink-200 shadow-md shadow-pink-500/5")
                              : (isDark ? "bg-[#140e10] opacity-60" : "bg-gray-50/50 opacity-60")
                          } flex flex-col justify-between transition-all relative overflow-hidden`}
                        >
                          {gw.enabled && (
                            <div className="absolute -top-10 -right-10 w-24 h-24 bg-pink-500/5 rounded-full blur-xl pointer-events-none" />
                          )}

                          <div>
                            {/* Header inside Card */}
                            <div className="flex items-center justify-between pb-4 border-b border-pink-500/10 mb-4">
                              <div className="flex items-center space-x-2.5">
                                <div className={`p-2 rounded-xl ${gw.enabled ? "bg-pink-500 text-white shadow-md shadow-pink-500/20" : "bg-gray-200 dark:bg-pink-950/40 text-gray-400 dark:text-pink-300/40"}`}>
                                  <CreditCard size={16} />
                                </div>
                                <div>
                                  <span className="text-[9px] font-bold text-gray-400 uppercase tracking-wider">
                                    {isDefault ? "Pasarela Oficial" : "Pasarela Personalizada"}
                                  </span>
                                  <h4 className={`text-sm font-bold ${textTitle}`}>
                                    {gw.name}
                                  </h4>
                                </div>
                              </div>

                              <button
                                type="button"
                                disabled={!canEdit}
                                onClick={() => handleUpdateGateway(gw.id, { enabled: !gw.enabled })}
                                className={`relative inline-flex h-6 w-11 shrink-0 cursor-pointer rounded-full border-2 border-transparent transition-colors duration-200 ease-in-out focus:outline-none ${
                                  gw.enabled ? "bg-pink-500" : "bg-gray-200 dark:bg-pink-950/60"
                                } ${!canEdit ? "opacity-50 cursor-not-allowed" : ""}`}
                              >
                                <span
                                  className={`pointer-events-none inline-block h-5 w-5 transform rounded-full bg-white shadow ring-0 transition duration-200 ease-in-out ${
                                    gw.enabled ? "translate-x-5" : "translate-x-0"
                                  }`}
                                />
                              </button>
                            </div>

                            {canEdit ? (
                              <div className="space-y-4">
                                <div>
                                  <label className="block text-[10px] font-bold uppercase text-gray-400 mb-1">Nombre del Método</label>
                                  <input
                                    type="text"
                                    value={gw.name}
                                    onChange={(e) => handleUpdateGateway(gw.id, { name: e.target.value })}
                                    className={inputStyle}
                                    placeholder="Ej. Banco Popular Dominicano"
                                  />
                                </div>

                                <div>
                                  <label className="block text-[10px] font-bold uppercase text-gray-400 mb-1">Enlace de Pago o Depósito (URL)</label>
                                  <input
                                    type="text"
                                    value={gw.link}
                                    onChange={(e) => handleUpdateGateway(gw.id, { link: e.target.value })}
                                    className={inputStyle}
                                    placeholder="https://..."
                                  />
                                </div>

                                <div>
                                  <label className="block text-[10px] font-bold uppercase text-gray-400 mb-1">Instrucciones Detalladas para Paciente</label>
                                  <textarea
                                    value={gw.instructions}
                                    onChange={(e) => handleUpdateGateway(gw.id, { instructions: e.target.value })}
                                    className={`${inputStyle} h-20 resize-none py-2 text-xs`}
                                    placeholder="Detalles de la cuenta corriente, titular y cómo enviar el comprobante..."
                                  />
                                </div>
                              </div>
                            ) : (
                              <div className="space-y-3 text-xs">
                                <div className="p-3 bg-rose-50/10 dark:bg-pink-950/10 rounded-xl border border-pink-500/5 space-y-1">
                                  <span className="text-[9px] font-bold text-pink-500 uppercase tracking-wider block">Instrucciones del Cliente:</span>
                                  <p className="text-gray-600 dark:text-pink-100 font-medium leading-relaxed italic">{gw.instructions}</p>
                                </div>

                                <div className="flex items-center justify-between text-[11px] text-gray-400 pt-1">
                                  <span>Enlace Asociado:</span>
                                  <a href={gw.link} target="_blank" rel="noopener noreferrer" className="font-mono text-pink-500 underline truncate max-w-[200px]" title={gw.link}>
                                    {gw.link}
                                  </a>
                                </div>
                              </div>
                            )}
                          </div>

                          {canEdit && !isDefault && (
                            <div className="mt-4 pt-4 border-t border-pink-500/5 flex justify-end">
                              <button
                                type="button"
                                onClick={() => handleDeleteGateway(gw.id)}
                                className="px-3 py-1.5 bg-rose-50 hover:bg-rose-100 dark:bg-rose-950/20 dark:hover:bg-rose-950/40 text-rose-600 dark:text-rose-400 rounded-xl text-[10px] font-bold uppercase tracking-wider transition-all flex items-center space-x-1"
                              >
                                <Trash2 size={12} />
                                <span>Eliminar</span>
                              </button>
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                </div>
              ) : (
                <>
                  {/* Left Form: Register Payments */}
                  <div className={`${bgCard} lg:col-span-1 h-fit`}>
                    <div className="flex items-center space-x-2 border-b pb-4 mb-5 border-pink-500/10">
                      <DollarSign className="text-pink-500" size={20} />
                      <h3 className={`text-lg font-bold ${textTitle}`}>{t.recordPayment}</h3>
                    </div>

                    <form onSubmit={handleRegisterPayment} className="space-y-4">
                      <div>
                        <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5">{t.selectPatient} *</label>
                        <select
                          value={newPayment.patientId}
                          onChange={e => setNewPayment(prev => ({ ...prev, patientId: e.target.value }))}
                          className={inputStyle}
                          required
                        >
                          <option value="">-- {lang === "es" ? "Seleccione un paciente" : "Select patient"} --</option>
                          {patients.map(p => (
                            <option key={p.id} value={p.id}>{p.name} (DNI: {p.dni})</option>
                          ))}
                        </select>
                      </div>

                      <div>
                        <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5">{t.treatmentType} *</label>
                        <select
                          value={newPayment.treatmentType}
                          onChange={e => setNewPayment(prev => ({ ...prev, treatmentType: e.target.value }))}
                          className={inputStyle}
                          required
                        >
                          <option value="">-- {lang === "es" ? "Seleccione tratamiento" : "Select treatment"} --</option>
                          <option value="Resina Estética">Resina Estética — $120</option>
                          <option value="Carillas de Disilicato">Carillas de Disilicato de Litio — $1,500</option>
                          <option value="Blanqueamiento Láser">Blanqueamiento Láser Dental — $300</option>
                          <option value="Endodoncia">Endodoncia / Tratamiento de Canal — $250</option>
                          <option value="Limpieza Profiláctica">Limpieza Profiláctica & Ultrasonido — $80</option>
                          <option value="Implante Titanio">Implante Dental de Titanio — $1,200</option>
                          <option value="Corona Zirconio">Corona de Zirconio Estética — $600</option>
                        </select>
                      </div>

                      <div className="grid grid-cols-2 gap-4">
                        <div>
                          <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5">{t.amountToPay} ($) *</label>
                          <input
                            type="number"
                            placeholder="1500"
                            value={newPayment.amountTotal || ""}
                            onChange={e => setNewPayment(prev => ({ ...prev, amountTotal: Number(e.target.value) }))}
                            className={inputStyle}
                            required
                          />
                        </div>
                        <div>
                          <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5">{t.amountPaid} ($)</label>
                          <input
                            type="number"
                            placeholder="1000"
                            value={newPayment.amountPaid || ""}
                            onChange={e => setNewPayment(prev => ({ ...prev, amountPaid: Number(e.target.value) }))}
                            className={inputStyle}
                          />
                        </div>
                      </div>

                      <div>
                        <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5">{t.paymentMethod} *</label>
                        <select
                          value={newPayment.paymentMethod}
                          onChange={e => setNewPayment(prev => ({ ...prev, paymentMethod: e.target.value as any }))}
                          className={inputStyle}
                        >
                          <option value="Cash">{t.cash}</option>
                          <option value="Card">{t.card}</option>
                          <option value="Transfer">{t.transfer}</option>
                        </select>
                      </div>

                      <button
                        type="submit"
                        className="w-full py-3 bg-gradient-to-r from-pink-400 to-rose-500 hover:from-pink-500 hover:to-rose-600 text-white font-bold rounded-xl text-xs uppercase tracking-widest shadow-md shadow-pink-500/10 transition-all flex items-center justify-center space-x-2"
                      >
                        <PlusCircle size={15} />
                        <span>{lang === "es" ? "Registrar Cobro" : "Register Billing"}</span>
                      </button>
                    </form>
                  </div>

                  {/* Right Columns: Stats and Ledger list */}
                  <div className="lg:col-span-2 space-y-6">
                    {/* Ledger metrics cards */}
                    <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                      <div className={`${bgCard} p-4 text-center border-l-4 border-l-pink-400`}>
                        <div className="text-gray-400 text-[10px] font-bold uppercase tracking-wider">{t.totalInvoiced}</div>
                        <div className={`text-2xl font-black ${textTitle} mt-1`}>${totalInvoicedAmt}</div>
                      </div>
                      <div className={`${bgCard} p-4 text-center border-l-4 border-l-emerald-400`}>
                        <div className="text-gray-400 text-[10px] font-bold uppercase tracking-wider">{t.totalCollected}</div>
                        <div className="text-2xl font-black text-emerald-500 mt-1">${totalCollectedAmt}</div>
                      </div>
                      <div className={`${bgCard} p-4 text-center border-l-4 border-l-amber-400`}>
                        <div className="text-gray-400 text-[10px] font-bold uppercase tracking-wider">{t.pendingReceivable}</div>
                        <div className="text-2xl font-black text-amber-500 mt-1">${pendingReceivableAmt}</div>
                      </div>
                    </div>

                    {/* Ledger ledger transactions */}
                    <div className={bgCard}>
                      <div className="flex items-center space-x-2 border-b pb-4 mb-4 border-pink-500/10">
                        <TrendingUp className="text-pink-500" size={20} />
                        <h3 className={`text-lg font-bold ${textTitle}`}>{t.billingHistory}</h3>
                      </div>

                      {payments.length === 0 ? (
                        <div className="text-center py-16 text-gray-400 space-y-3">
                          <DollarSign className="mx-auto text-gray-300" size={48} />
                          <p className="text-sm font-medium">{t.noPayments}</p>
                        </div>
                      ) : (
                        <div className="space-y-3 max-h-[380px] overflow-y-auto pr-2 custom-scrollbar">
                          {payments.map(pay => {
                            return (
                              <div
                                key={pay.id}
                                className={`p-4 rounded-xl border ${borderCol} ${isDark ? "bg-[#1f1618]" : "bg-rose-50/10"} flex items-center justify-between gap-4`}
                              >
                                <div className="space-y-1">
                                  <h4 className={`text-sm font-bold ${textTitle}`}>{pay.patientName}</h4>
                                  <p className="text-xs text-gray-400 font-bold uppercase tracking-wider">{pay.treatmentType}</p>
                                  <div className="flex items-center space-x-3 text-[10px] font-mono text-gray-400">
                                    <span className="bg-pink-500/10 text-pink-500 px-2 py-0.5 rounded uppercase font-bold">{pay.paymentMethod}</span>
                                    <span>{pay.date}</span>
                                  </div>
                                </div>

                                <div className="text-right flex items-center space-x-4">
                                  <div>
                                    <div className={`text-sm font-black ${textTitle}`}>${pay.amountPaid} <span className="text-[10px] text-gray-400 font-semibold">/ ${pay.amountTotal}</span></div>
                                    {pay.amountPending > 0 && (
                                      <div className="text-[10px] text-amber-500 font-bold">{lang === "es" ? "Debe" : "Owes"} ${pay.amountPending}</div>
                                    )}
                                  </div>

                                  <button
                                    onClick={() => setSelectedInvoice(pay)}
                                    className={`p-2 rounded-xl border ${borderCol} hover:bg-pink-500/10 text-pink-500 transition-all`}
                                    title={t.invoice}
                                  >
                                    <Printer size={13} />
                                  </button>
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      )}
                    </div>
                  </div>
                </>
              )}
            </motion.div>
          )}

          {/* TAB 5: ASISTENTE CLINICO IA GEMINI */}
          {activeTab === "ai" && (
            <motion.div
              key="ai-view"
              initial={{ opacity: 0, y: 15 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -15 }}
              className="space-y-6"
            >
              <div className={`${bgCard} max-w-4xl mx-auto space-y-6`}>
                <div className="flex items-center space-x-3 border-b border-pink-500/5 pb-4">
                  <div className="p-2.5 bg-gradient-to-tr from-pink-400 to-rose-500 text-white rounded-xl shadow-md">
                    <Sparkles size={20} />
                  </div>
                  <div>
                    <h3 className={`text-lg font-bold ${textTitle}`}>{t.aiTitle}</h3>
                    <p className={`text-xs ${textSub}`}>{t.aiSubtitle}</p>
                  </div>
                </div>

                <div className="space-y-4">
                  <label className="block text-xs font-bold uppercase text-gray-500">{lang === "es" ? "Diagnóstico Asistido o Cuadro Clínico:" : "Clinical Findings Input:"}</label>
                  <textarea
                    rows={4}
                    placeholder={t.aiPromptPlaceholder}
                    value={aiPrompt}
                    onChange={e => setAiPrompt(e.target.value)}
                    className={inputStyle}
                  />

                  {/* Predefined prompt helpers */}
                  <div className="flex flex-wrap gap-2 pt-1">
                    {[
                      { es: "Dolor punzante frío diente 16", en: "Sharp cold pain tooth 16" },
                      { es: "Gingivitis sangrante generalizada", en: "General bleeding gums gingivitis" },
                      { es: "Corona desadaptada molar 46 con sensibilidad", en: "Molar 46 loose crown with sensitivity" }
                    ].map((helper, i) => (
                      <button
                        key={i}
                        type="button"
                        onClick={() => setAiPrompt(lang === "es" ? helper.es : helper.en)}
                        className="px-3 py-1.5 border border-pink-500/5 bg-pink-500/5 text-pink-500 hover:bg-pink-500/10 text-[10px] font-bold rounded-lg transition-all uppercase tracking-wider"
                      >
                        {lang === "es" ? helper.es : helper.en}
                      </button>
                    ))}
                  </div>
                </div>

                <button
                  onClick={handleAiAnalyze}
                  disabled={isAiLoading || !aiPrompt.trim()}
                  className="w-full py-3 bg-gradient-to-r from-pink-400 to-rose-500 hover:from-pink-500 hover:to-rose-600 disabled:from-gray-400 disabled:to-gray-500 text-white font-bold rounded-xl text-xs uppercase tracking-widest shadow-md shadow-pink-500/10 transition-all flex items-center justify-center space-x-2"
                >
                  <RefreshCw size={15} className={isAiLoading ? "animate-spin" : ""} />
                  <span>{isAiLoading ? t.aiAnalyzing : t.aiAnalyze}</span>
                </button>

                {/* Response result rendering */}
                {aiResponse && (
                  <motion.div
                    initial={{ opacity: 0, y: 10 }}
                    animate={{ opacity: 1, y: 0 }}
                    className={`p-6 rounded-2xl border border-pink-500/10 ${isDark ? "bg-[#1f1618]" : "bg-rose-50/20"} space-y-4`}
                  >
                    <div className="flex items-center justify-between border-b border-pink-500/5 pb-3">
                      <div className="flex items-center space-x-2 text-xs font-bold text-pink-500 font-mono uppercase tracking-wider">
                        <Sparkles size={14} />
                        <span>{t.aiResult}</span>
                      </div>
                      <span className="text-[9px] bg-emerald-500/10 text-emerald-500 border border-emerald-500/20 font-bold uppercase px-2 py-0.5 rounded-full tracking-widest">
                        Gemini-2.5-Flash Active
                      </span>
                    </div>

                    <p className="text-xs text-gray-400 italic">
                      {t.aiResultIntro}
                    </p>

                    <div className="prose prose-pink max-w-none text-xs md:text-sm text-gray-500 dark:text-pink-100/90 leading-relaxed font-sans whitespace-pre-wrap space-y-4">
                      {aiResponse}
                    </div>
                  </motion.div>
                )}
              </div>
            </motion.div>
          )}

          {/* TAB 5.5: DISPONIBILIDAD DE DOCTORES */}
          {activeTab === "availability" && (
            <motion.div
              key="availability-view"
              initial={{ opacity: 0, y: 15 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -15 }}
              className="grid grid-cols-1 lg:grid-cols-3 gap-8"
            >
              {/* Left Column: Form to configure availability */}
              <div className={`${bgCard} lg:col-span-1 h-fit space-y-5`}>
                <div className="flex items-center space-x-2 border-b pb-4 border-pink-500/10">
                  <Clock className="text-pink-500" size={20} />
                  <h3 className={`text-lg font-bold ${textTitle}`}>
                    {lang === "es" ? "Programar Disponibilidad" : "Schedule Availability"}
                  </h3>
                </div>

                <form onSubmit={handleSaveAvailability} className="space-y-4">
                  {/* Doctor selection */}
                  <div>
                    <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5">
                      {lang === "es" ? "Doctor / Especialista" : "Doctor / Specialist"}
                    </label>
                    {currentAdmin?.role === "master" ? (
                      <select
                        value={availDoctor}
                        onChange={e => setAvailDoctor(e.target.value)}
                        className={inputStyle}
                      >
                        <option value="">-- {lang === "es" ? "Seleccione Doctor" : "Select Doctor"} --</option>
                        {Array.from(new Set([
                          "Dra. Diana Rojas",
                          "Dr. Marcos Soler",
                          "Dra. Lucía Santos",
                          ...admins.map(a => a.name)
                        ])).map(name => (
                          <option key={name} value={name}>{name}</option>
                        ))}
                      </select>
                    ) : (
                      <div className="flex items-center space-x-2 p-3 rounded-xl bg-pink-500/5 border border-pink-500/10 text-xs font-semibold text-pink-500">
                        <Shield size={14} />
                        <span>{availDoctor || currentAdmin?.name}</span>
                        <span className="text-[9px] uppercase bg-pink-500/10 text-pink-500 px-1.5 py-0.5 rounded ml-auto">
                          {lang === "es" ? "Bloqueado" : "Locked"}
                        </span>
                      </div>
                    )}
                  </div>

                  {/* Date selection */}
                  <div>
                    <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5">
                      {lang === "es" ? "Fecha" : "Date"}
                    </label>
                    <input
                      type="date"
                      required
                      min={new Date().toISOString().split("T")[0]}
                      value={availDate}
                      onChange={e => setAvailDate(e.target.value)}
                      className={inputStyle}
                    />
                  </div>

                  {/* Hourly slots toggles */}
                  <div>
                    <div className="flex items-center justify-between mb-1.5">
                      <label className="block text-xs font-bold uppercase text-gray-500">
                        {lang === "es" ? "Bloques de Horas" : "Hour Slots"}
                      </label>
                      <span className="text-[10px] text-pink-500 font-bold">
                        {availSlots.length} {lang === "es" ? "seleccionados" : "selected"}
                      </span>
                    </div>

                    {/* Presets */}
                    <div className="flex flex-wrap gap-1 mb-3">
                      <button
                        type="button"
                        onClick={() => setAvailSlots(["08:00", "09:00", "10:00", "11:00", "12:00"])}
                        className="text-[9px] font-bold bg-pink-500/5 text-pink-500 border border-pink-500/10 px-2.5 py-1 rounded hover:bg-pink-500/15 cursor-pointer"
                      >
                        🌅 {lang === "es" ? "Mañana" : "Morning"}
                      </button>
                      <button
                        type="button"
                        onClick={() => setAvailSlots(["14:00", "15:00", "16:00", "17:00", "18:00"])}
                        className="text-[9px] font-bold bg-pink-500/5 text-pink-500 border border-pink-500/10 px-2.5 py-1 rounded hover:bg-pink-500/15 cursor-pointer"
                      >
                        🌇 {lang === "es" ? "Tarde" : "Afternoon"}
                      </button>
                      <button
                        type="button"
                        onClick={() => setAvailSlots(["08:00", "09:00", "10:00", "11:00", "12:00", "14:00", "15:00", "16:00", "17:00", "18:00"])}
                        className="text-[9px] font-bold bg-pink-500/5 text-pink-500 border border-pink-500/10 px-2.5 py-1 rounded hover:bg-pink-500/15 cursor-pointer"
                      >
                        📅 {lang === "es" ? "Todo el Día" : "Full Day"}
                      </button>
                      <button
                        type="button"
                        onClick={() => setAvailSlots([])}
                        className="text-[9px] font-bold bg-red-500/5 text-red-500 border border-red-500/10 px-2.5 py-1 rounded hover:bg-red-500/15 cursor-pointer ml-auto"
                      >
                        🗑️ {lang === "es" ? "Limpiar" : "Clear"}
                      </button>
                    </div>

                    {/* Slots Grid */}
                    <div className="grid grid-cols-3 gap-2 p-3 bg-pink-500/5 dark:bg-black/20 rounded-2xl border border-pink-500/10">
                      {[
                        "08:00", "09:00", "10:00", "11:00", "12:00",
                        "13:00", "14:00", "15:00", "16:00", "17:00", "18:00"
                      ].map(time => {
                        const isSelected = availSlots.includes(time);
                        return (
                          <button
                            key={time}
                            type="button"
                            onClick={() => {
                              if (isSelected) {
                                setAvailSlots(availSlots.filter(s => s !== time));
                              } else {
                                setAvailSlots([...availSlots, time].sort());
                              }
                            }}
                            className={`py-1.5 text-xs font-mono font-bold rounded-lg transition-all cursor-pointer ${
                              isSelected
                                ? "bg-gradient-to-r from-pink-400 to-rose-500 text-white shadow shadow-pink-500/10 scale-[1.03]"
                                : "bg-white dark:bg-[#1f1618] text-rose-800 dark:text-pink-300 border border-pink-500/10 hover:border-pink-500/40"
                            }`}
                          >
                            {time}
                          </button>
                        );
                      })}
                    </div>
                  </div>

                  {/* Submit button */}
                  <button
                    type="submit"
                    className="w-full py-3 bg-gradient-to-r from-pink-400 to-rose-500 hover:from-pink-500 hover:to-rose-600 text-white font-bold rounded-xl text-xs uppercase tracking-wider shadow-md shadow-pink-500/10 transition-all flex items-center justify-center space-x-2 cursor-pointer"
                  >
                    <PlusCircle size={14} />
                    <span>{lang === "es" ? "Guardar Disponibilidad" : "Save Availability"}</span>
                  </button>
                </form>
              </div>

              {/* Right Column: List of scheduled availabilities */}
              <div className={`${bgCard} lg:col-span-2 flex flex-col space-y-4`}>
                <div className="flex items-center justify-between border-b pb-4 border-pink-500/10">
                  <div className="flex items-center space-x-2">
                    <Calendar className="text-pink-500" size={20} />
                    <h3 className={`text-lg font-bold ${textTitle}`}>
                      {lang === "es" ? "Calendario de Disponibilidad Clínica" : "Clinical Availability Calendar"}
                    </h3>
                  </div>
                  <span className="text-xs bg-pink-100 dark:bg-pink-950/40 text-pink-600 dark:text-pink-400 font-bold px-3 py-1 rounded-full border border-pink-200/40 dark:border-pink-500/10">
                    {availabilities.length} {lang === "es" ? "Días Programados" : "Days Scheduled"}
                  </span>
                </div>

                {availabilities.length === 0 ? (
                  <div className="text-center py-16 text-gray-400 space-y-3 flex-grow flex flex-col items-center justify-center">
                    <Clock className="mx-auto text-gray-300" size={48} />
                    <p className="text-sm font-medium">
                      {lang === "es" ? "No hay disponibilidades programadas." : "No availabilities scheduled yet."}
                    </p>
                    <p className="text-xs text-gray-400 max-w-sm">
                      {lang === "es"
                        ? "Use el formulario de la izquierda para definir qué días y horas atienden los doctores."
                        : "Use the form on the left to define which days and times doctors are attending."}
                    </p>
                  </div>
                ) : (
                  <div className="space-y-4 overflow-y-auto max-h-[580px] pr-2 custom-scrollbar flex-grow">
                    {[...availabilities]
                      .sort((a, b) => a.date.localeCompare(b.date))
                      .map(avail => {
                        return (
                          <div
                            key={avail.id}
                            className={`p-5 rounded-2xl border ${borderCol} ${
                              isDark ? "bg-[#1f1618]" : "bg-rose-50/10"
                            } space-y-3`}
                          >
                            <div className="flex items-center justify-between border-b border-pink-500/5 pb-2">
                              <div>
                                <h4 className={`text-sm font-bold ${textTitle}`}>{avail.doctorName}</h4>
                                <div className="text-[10px] text-pink-500 font-mono font-bold flex items-center space-x-1 mt-0.5">
                                  <Calendar size={10} />
                                  <span>{avail.date}</span>
                                </div>
                              </div>

                              <button
                                type="button"
                                onClick={() => handleDeleteAvailability(avail.id)}
                                className="p-1.5 text-rose-500 hover:text-rose-600 hover:bg-rose-500/5 rounded-lg transition-colors cursor-pointer"
                                title={lang === "es" ? "Eliminar Día" : "Delete Day"}
                              >
                                <Trash2 size={14} />
                              </button>
                            </div>

                            {/* List of slots inside this day */}
                            <div className="flex flex-wrap gap-2">
                              {avail.slots.map(slot => {
                                // Check if this slot has a scheduled booking
                                const activeBooking = appointments.find(
                                  app =>
                                    app.dentistName.toLowerCase() === avail.doctorName.toLowerCase() &&
                                    app.date === avail.date &&
                                    app.time === slot &&
                                    app.status === "Scheduled"
                                );

                                if (activeBooking) {
                                  return (
                                    <div
                                      key={slot}
                                      className="px-2.5 py-1 text-xs font-semibold rounded-lg bg-rose-500/10 text-rose-600 dark:text-pink-300 border border-rose-500/20 flex items-center space-x-1.5"
                                      title={`Cita de: ${activeBooking.patientName} (${activeBooking.treatmentType})`}
                                    >
                                      <span className="w-1.5 h-1.5 rounded-full bg-rose-500 animate-pulse" />
                                      <span>{slot}</span>
                                      <span className="text-[9px] font-normal tracking-wide text-rose-500 dark:text-pink-400 max-w-[80px] truncate">
                                        ({activeBooking.patientName.split(" ")[0]})
                                      </span>
                                    </div>
                                  );
                                } else {
                                  return (
                                    <div
                                      key={slot}
                                      className="px-2.5 py-1 text-xs font-semibold rounded-lg bg-emerald-500/10 text-emerald-600 dark:text-emerald-400 border border-emerald-500/20 flex items-center space-x-1"
                                    >
                                      <span className="w-1.5 h-1.5 rounded-full bg-emerald-500" />
                                      <span>{slot}</span>
                                    </div>
                                  );
                                }
                              })}
                            </div>
                          </div>
                        );
                      })}
                  </div>
                )}
              </div>
            </motion.div>
          )}

          {/* TAB 6: SINCRONIZACIÓN GOOGLE SHEETS */}
          {activeTab === "sync" && (
            <motion.div
              key="sync-view"
              initial={{ opacity: 0, y: 15 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -15 }}
              className="grid grid-cols-1 lg:grid-cols-3 gap-8"
            >
              {/* Left Column: API parameters configuration & Admin Accounts */}
              <div className="lg:col-span-1 flex flex-col gap-6">
                
                {/* Google Sheets Sync Card */}
                <div className={`${bgCard} space-y-6`}>
                  <div className="flex items-center justify-between border-b pb-4 border-pink-500/10">
                    <div className="flex items-center space-x-2">
                      <Database className="text-pink-500" size={20} />
                      <h3 className={`text-lg font-bold ${textTitle}`}>{t.syncCenter}</h3>
                    </div>
                    <span className="text-[10px] bg-emerald-500/10 text-emerald-500 font-mono font-extrabold uppercase px-2 py-0.5 rounded-full border border-emerald-500/15 animate-pulse">
                      PROD LIVE
                    </span>
                  </div>

                  {/* Direct Link to the Google Spreadsheet */}
                  <div className="p-3.5 rounded-2xl bg-emerald-500/5 border border-emerald-500/15 space-y-2">
                    <div className="flex items-center justify-between">
                      <span className="text-[10px] font-bold uppercase text-emerald-600 dark:text-emerald-400">
                        {lang === "es" ? "Base de Datos en Tiempo Real" : "Real-time Database"}
                      </span>
                      <span className="w-2 h-2 bg-emerald-500 rounded-full animate-ping" />
                    </div>
                    <p className="text-[11px] text-gray-500 dark:text-pink-300/60 leading-normal">
                      {lang === "es" 
                        ? "DianaSRL está sincronizado directamente con la siguiente hoja oficial de Google Sheets:" 
                        : "DianaSRL is directly synchronized with the following official Google Sheet:"}
                    </p>
                    <a 
                      href={spreadsheetUrl} 
                      target="_blank" 
                      rel="noopener noreferrer"
                      className="inline-flex items-center space-x-1.5 text-xs font-bold text-emerald-600 hover:text-emerald-700 dark:text-emerald-400 dark:hover:text-emerald-300 underline underline-offset-4 decoration-emerald-500/30 transition-all font-mono break-all"
                    >
                      <span>{lang === "es" ? "Abrir Google Sheets Live ↗" : "Open Live Google Sheets ↗"}</span>
                    </a>
                  </div>

                  <div className="space-y-4">
                    <div className="flex items-center justify-between">
                      <span className="text-xs font-bold uppercase text-gray-500">{t.syncStatus}:</span>
                      <button
                        type="button"
                        onClick={() => {
                          const next = !liveSync;
                          setLiveSync(next);
                          localStorage.setItem("dianasrl_live_sync", JSON.stringify(next));
                          triggerToast(next ? t.syncActive : t.syncDisabled);
                        }}
                        className={`px-3 py-1 rounded-full text-[10px] font-extrabold uppercase border transition-all ${
                          liveSync 
                            ? "bg-emerald-500/10 text-emerald-500 border-emerald-500/20" 
                            : "bg-gray-100 text-gray-500 border-gray-300 dark:bg-gray-800 dark:text-gray-400 dark:border-gray-700"
                        }`}
                      >
                        {liveSync ? t.syncActive : t.syncDisabled}
                      </button>
                    </div>

                    <div>
                      <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5">
                        {lang === "es" ? "URL de la Planilla de Google Sheets" : "Google Sheets Spreadsheet URL"}
                      </label>
                      <input
                        type="text"
                        placeholder="https://docs.google.com/spreadsheets/d/.../edit"
                        value={spreadsheetUrl}
                        onChange={e => {
                          setSpreadsheetUrl(e.target.value);
                          localStorage.setItem("dianasrl_spreadsheet_url", e.target.value);
                        }}
                        className={inputStyle}
                      />
                    </div>

                    <div>
                      <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5">{t.googleSheetsUrl}</label>
                      <input
                        type="text"
                        placeholder="https://script.google.com/macros/s/.../exec"
                        value={scriptUrl}
                        onChange={e => {
                          setScriptUrl(e.target.value);
                          localStorage.setItem("dianasrl_script_url", e.target.value);
                        }}
                        className={inputStyle}
                      />
                    </div>

                    <div className="flex flex-col gap-2.5">
                      <button
                        type="button"
                        onClick={handleTestSync}
                        disabled={isTestingSync || !scriptUrl}
                        className="w-full py-2.5 bg-[#1e1416] border border-pink-500/20 hover:border-pink-500 text-white font-bold rounded-xl text-xs uppercase tracking-wider transition-all flex items-center justify-center space-x-2"
                      >
                        <RefreshCw size={14} className={isTestingSync ? "animate-spin" : ""} />
                        <span>{isTestingSync ? t.testing : t.testSync}</span>
                      </button>

                      <button
                        type="button"
                        onClick={handleImportAllFromSheets}
                        disabled={isImportingSync || !scriptUrl}
                        className="w-full py-2.5 bg-gradient-to-r from-emerald-400 to-teal-500 hover:from-emerald-500 hover:to-teal-600 disabled:from-gray-400 disabled:to-gray-500 text-white font-bold rounded-xl text-xs uppercase tracking-wider transition-all flex items-center justify-center space-x-2 shadow-md shadow-emerald-500/10"
                      >
                        <Download size={14} className={isImportingSync ? "animate-spin" : ""} />
                        <span>{isImportingSync ? "Cargando..." : "Cargar desde Google Sheets"}</span>
                      </button>

                      <button
                        type="button"
                        onClick={handleExportAllToSheets}
                        disabled={isExportingSync || !scriptUrl}
                        className="w-full py-2.5 bg-gradient-to-r from-pink-400 to-rose-500 hover:from-pink-500 hover:to-rose-600 disabled:from-gray-400 disabled:to-gray-500 text-white font-bold rounded-xl text-xs uppercase tracking-wider transition-all flex items-center justify-center space-x-2 shadow-md shadow-pink-500/10"
                      >
                        <Database size={14} className={isExportingSync ? "animate-spin" : ""} />
                        <span>{isExportingSync ? "Subiendo..." : "Subir todo a Google Sheets"}</span>
                      </button>
                    </div>
                  </div>

                  {/* Portal Links Direct Access Card */}
                  <div className={`${bgCard} space-y-4`}>
                    <div className="flex items-center justify-between border-b pb-3 border-pink-500/10">
                      <div className="flex items-center space-x-2">
                        <Layers className="text-pink-500" size={18} />
                        <h4 className={`text-sm font-bold ${textTitle}`}>
                          {lang === "es" ? "Enlaces de Portales de Acceso" : "Portal Access Links"}
                        </h4>
                      </div>
                      <span className="text-[9px] bg-pink-500/10 text-pink-500 font-mono font-bold uppercase px-2 py-0.5 rounded-full border border-pink-500/20">
                        {lang === "es" ? "Dividido" : "Divided"}
                      </span>
                    </div>

                    <p className="text-[11px] text-gray-500 dark:text-pink-300/60 leading-normal">
                      {lang === "es" 
                        ? "Utilice estos enlaces únicos para ingresar directamente a la vista de cliente (paciente) o a la vista clínica de administración:"
                        : "Use these unique links to access directly either the client (patient) view or the clinical admin view:"}
                    </p>

                    <div className="space-y-3">
                      {/* Patient Portal Link */}
                      <div className="p-3 rounded-xl bg-pink-500/5 border border-pink-500/10 space-y-1.5">
                        <div className="flex items-center justify-between">
                          <span className="text-[10px] font-bold text-pink-500 uppercase tracking-wider flex items-center gap-1">
                            <Heart size={10} className="fill-current" />
                            {lang === "es" ? "Portal del Paciente (Cliente)" : "Patient Portal (Client)"}
                          </span>
                          <button
                            type="button"
                            onClick={() => {
                              const link = `${window.location.origin}${window.location.pathname}?portal=patient`;
                              navigator.clipboard.writeText(link);
                              triggerToast(lang === "es" ? "¡Enlace del Portal de Pacientes copiado!" : "Patient Portal link copied!", "success");
                            }}
                            className="text-[10px] font-bold text-pink-500 hover:text-pink-600 dark:hover:text-pink-400 underline cursor-pointer"
                          >
                            {lang === "es" ? "Copiar" : "Copy"}
                          </button>
                        </div>
                        <div className="text-[10.5px] font-mono text-gray-400 break-all select-all bg-black/10 dark:bg-black/30 p-1.5 rounded-lg border border-pink-500/5">
                          {`${window.location.origin}${window.location.pathname}?portal=patient`}
                        </div>
                      </div>

                      {/* Admin Portal Link */}
                      <div className="p-3 rounded-xl bg-rose-500/5 border border-rose-500/10 space-y-1.5">
                        <div className="flex items-center justify-between">
                          <span className="text-[10px] font-bold text-rose-600 dark:text-pink-300 uppercase tracking-wider flex items-center gap-1">
                            <Users size={10} />
                            {lang === "es" ? "Portal de Administración (Clínico)" : "Admin Portal (Clinical)"}
                          </span>
                          <button
                            type="button"
                            onClick={() => {
                              const link = `${window.location.origin}${window.location.pathname}?portal=admin`;
                              navigator.clipboard.writeText(link);
                              triggerToast(lang === "es" ? "¡Enlace del Portal de Administración copiado!" : "Admin Portal link copied!", "success");
                            }}
                            className="text-[10px] font-bold text-rose-600 dark:text-pink-300 hover:text-rose-700 dark:hover:text-pink-200 underline cursor-pointer"
                          >
                            {lang === "es" ? "Copiar" : "Copy"}
                          </button>
                        </div>
                        <div className="text-[10.5px] font-mono text-gray-400 break-all select-all bg-black/10 dark:bg-black/30 p-1.5 rounded-lg border border-pink-500/5">
                          {`${window.location.origin}${window.location.pathname}?portal=admin`}
                        </div>
                      </div>
                    </div>
                  </div>

                  {/* Google Apps Script Integration Helper Tutorial */}
                  <div className={`p-4 rounded-2xl border ${borderCol} ${isDark ? "bg-[#1f1618]" : "bg-rose-50/20"} space-y-3`}>
                    <div className="flex items-center justify-between flex-wrap gap-2">
                      <h4 className="text-xs font-bold text-pink-500 flex items-center space-x-1.5 uppercase tracking-wider">
                        <Info size={14} />
                        <span>Tutorial de Sincronización Bidireccional</span>
                      </h4>
                      <button
                        type="button"
                        onClick={() => {
                          navigator.clipboard.writeText(APPS_SCRIPT_CODE);
                          setCopiedScript(true);
                          triggerToast(
                            lang === "es"
                              ? "¡Código de Apps Script copiado al portapapeles!"
                              : "Apps Script code copied to clipboard!",
                            "success"
                          );
                          setTimeout(() => setCopiedScript(false), 2000);
                        }}
                        className={`flex items-center space-x-1 text-[10px] font-bold px-2.5 py-1 rounded-lg border transition-all cursor-pointer ${
                          copiedScript
                            ? "bg-emerald-500/10 border-emerald-500/30 text-emerald-500"
                            : "bg-pink-500/10 border-pink-500/20 text-pink-500 hover:bg-pink-500/20"
                        }`}
                      >
                        {copiedScript ? <Check size={12} /> : <Copy size={12} />}
                        <span>
                          {copiedScript
                            ? (lang === "es" ? "¡Copiado!" : "Copied!")
                            : (lang === "es" ? "Copiar Código" : "Copy Code")}
                        </span>
                      </button>
                    </div>
                    <p className="text-[11px] text-gray-500 dark:text-pink-300/60 leading-relaxed">
                      Este sistema permite tanto subir todos sus datos locales como descargarlos de Google Sheets. Para habilitar ambas funciones, reemplace todo el código en el editor de Apps Script (Menú <strong>Extensiones &gt; Apps Script</strong>) de su hoja con este código:
                    </p>
                    <pre className="text-[9px] p-3 rounded-lg bg-black/80 text-emerald-400 overflow-x-auto font-mono max-h-48 leading-normal border border-white/5">
{APPS_SCRIPT_CODE}
                    </pre>
                  </div>
                </div>

                {/* Admin Accounts Management Card */}
                {currentAdmin && (
                  <div className={bgCard}>
                    <div className="flex items-center justify-between border-b pb-4 mb-4 border-pink-500/10">
                      <div className="flex items-center space-x-2">
                        <Shield className="text-pink-500" size={18} />
                        <h3 className={`text-sm font-bold ${textTitle}`}>
                          {currentAdmin.role === "master" ? "Alta de Administradores" : "Cuentas Clínicas"}
                        </h3>
                      </div>
                      <span className="text-[9px] bg-pink-500/10 text-pink-500 px-2 py-0.5 rounded-full font-mono uppercase font-bold">
                        {currentAdmin.role === "master" ? "Master Account" : "Estándar"}
                      </span>
                    </div>

                    {currentAdmin.role === "master" ? (
                      <div className="space-y-4">
                        <p className="text-[11px] text-gray-500 dark:text-pink-300/60 leading-relaxed">
                          Como Fundadora Master de DianaSRL, usted cuenta con la exclusividad para añadir nuevos administradores de forma segura en la base de datos local.
                        </p>

                        <form onSubmit={handleRegisterAdmin} className="space-y-3 pt-1">
                          <div>
                            <label className="block text-[10px] font-bold uppercase text-gray-400 dark:text-pink-300/50 mb-1">Nombre del Administrador *</label>
                            <input
                              type="text"
                              required
                              placeholder="Ej. Dr. Marcos Soler"
                              value={newAdminForm.name}
                              onChange={e => setNewAdminForm({ ...newAdminForm, name: e.target.value })}
                              className={inputStyle}
                            />
                          </div>

                          <div>
                            <label className="block text-[10px] font-bold uppercase text-gray-400 dark:text-pink-300/50 mb-1">Usuario de Acceso *</label>
                            <input
                              type="text"
                              required
                              placeholder="Ej. msoler"
                              value={newAdminForm.username}
                              onChange={e => setNewAdminForm({ ...newAdminForm, username: e.target.value })}
                              className={inputStyle}
                            />
                          </div>

                          <div>
                            <label className="block text-[10px] font-bold uppercase text-gray-400 dark:text-pink-300/50 mb-1">Contraseña de Acceso *</label>
                            <input
                              type="password"
                              required
                              placeholder="Mínimo 4 caracteres"
                              value={newAdminForm.password}
                              onChange={e => setNewAdminForm({ ...newAdminForm, password: e.target.value })}
                              className={inputStyle}
                            />
                          </div>

                          <button
                            type="submit"
                            className="w-full py-2 bg-gradient-to-r from-pink-400 to-rose-500 hover:from-pink-500 hover:to-rose-600 text-white font-bold rounded-xl text-xs uppercase tracking-wider transition-all flex items-center justify-center space-x-1.5 shadow-md shadow-pink-500/10 hover:scale-[1.01]"
                          >
                            <Plus size={13} />
                            <span>Crear Nuevo Admin</span>
                          </button>
                        </form>
                      </div>
                    ) : (
                      <div className="space-y-3">
                        <div className="p-3 bg-rose-500/5 border border-rose-500/10 rounded-xl">
                          <p className="text-xs text-rose-500 font-semibold leading-relaxed">
                            Acceso de registro deshabilitado.
                          </p>
                          <p className="text-[11px] text-gray-500 dark:text-pink-300/50 leading-relaxed mt-1">
                            Su cuenta <strong>@{currentAdmin.username}</strong> tiene rol de Administrador Estándar. Solo el fundador Master (Dra. Diana Rojas) puede añadir otros administradores.
                          </p>
                        </div>
                      </div>
                    )}

                    {/* Active Admins List (always viewable) */}
                    <div className="mt-5 pt-4 border-t border-pink-500/5">
                      <h4 className="text-[10px] font-bold uppercase text-gray-400 mb-3 tracking-wider">Administradores Clínicos Activos:</h4>
                      <div className="space-y-2 max-h-48 overflow-y-auto pr-1">
                        {admins.map(adm => (
                          <div 
                            key={adm.id} 
                            className={`p-2.5 rounded-xl border ${borderCol} ${isDark ? "bg-[#1f1618]" : "bg-rose-50/10"} flex items-center justify-between text-xs`}
                          >
                            <div className="space-y-0.5">
                              <div className={`font-bold ${textTitle}`}>{adm.name}</div>
                              <div className="text-[10px] text-gray-400 font-mono">@{adm.username}</div>
                            </div>
                            <span className={`text-[9px] font-extrabold uppercase px-2 py-0.5 rounded-full font-mono ${
                              adm.role === "master" 
                                ? "bg-amber-500/10 text-amber-500 border border-amber-500/20" 
                                : "bg-pink-500/10 text-pink-500 border border-pink-500/20"
                            }`}>
                              {adm.role === "master" ? "Master" : "Admin"}
                            </span>
                          </div>
                        ))}
                      </div>
                    </div>
                  </div>
                )}
              </div>

              {/* Right Column: Ledger Log Activity */}
              <div className={`${bgCard} lg:col-span-2 flex flex-col`}>
                <div className="flex items-center justify-between border-b pb-4 mb-4 border-pink-500/10">
                  <div className="flex items-center space-x-2">
                    <Database className="text-pink-500" size={20} />
                    <h3 className={`text-lg font-bold ${textTitle}`}>{t.syncLogs}</h3>
                  </div>
                  <button
                    onClick={() => {
                      const clearedLogs = [{ id: "l0", timestamp: new Date().toISOString(), action: "Logs Cleared", status: "Success" as const, message: "Registro de actividad vaciado." }];
                      setSyncLogs(clearedLogs);
                      localStorage.setItem("dianasrl_sync_logs", JSON.stringify(clearedLogs));
                    }}
                    className="text-[10px] text-pink-500 hover:text-pink-600 font-bold uppercase transition-colors"
                  >
                    Vaciarlas
                  </button>
                </div>

                <p className="text-xs text-gray-400 italic mb-4">
                  {t.syncNotice}
                </p>

                {/* Sync log messages block */}
                <div className="space-y-2 flex-grow overflow-y-auto max-h-[460px] pr-2 custom-scrollbar">
                  {syncLogs.map(log => {
                    const statusColors = {
                      Success: "text-emerald-500 bg-emerald-500/10 border-emerald-500/20",
                      Error: "text-red-500 bg-red-500/10 border-red-500/20",
                      Pending: "text-amber-500 bg-amber-500/10 border-amber-500/20 animate-pulse"
                    };

                    return (
                      <div
                        key={log.id}
                        className={`p-3 rounded-lg border ${borderCol} flex flex-col md:flex-row justify-between gap-2 text-xs ${
                          isDark ? "bg-[#1f1618]" : "bg-rose-50/10"
                        }`}
                      >
                        <div className="space-y-1">
                          <div className="flex items-center space-x-2">
                            <span className={`px-2 py-0.5 rounded text-[9px] font-extrabold uppercase border ${statusColors[log.status]}`}>
                              {log.status}
                            </span>
                            <span className={`font-bold ${textTitle}`}>{log.action}</span>
                          </div>
                          <p className="text-gray-500 dark:text-pink-300/60 font-medium">{log.message}</p>
                        </div>

                        <span className="text-[10px] text-gray-400 shrink-0 font-mono font-bold">
                          {log.timestamp.split("T")[1]?.slice(0, 8) || log.timestamp}
                        </span>
                      </div>
                    );
                  })}
                </div>
              </div>
            </motion.div>
          )}

          {activeTab === "master-config" && currentAdmin?.role === "master" && (
            <motion.div
              initial={{ opacity: 0, y: 15 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: 15 }}
              className="grid grid-cols-1 lg:grid-cols-5 gap-6"
            >
              {/* BRANDING & CONTROLS SECTION */}
              <div className="lg:col-span-3 space-y-6">
                {/* Brand Identity Card */}
                <div className={bgCard}>
                  <div className="flex items-center space-x-2.5 border-b pb-4 mb-5 border-pink-500/10">
                    <Palette className={activeThemeObj.accentText} size={22} />
                    <div>
                      <h3 className={`text-lg font-bold ${textTitle}`}>
                        {lang === "es" ? "Identidad de la Clínica" : "Clinic Brand Identity"}
                      </h3>
                      <p className="text-[10px] text-gray-400 font-medium">
                        {lang === "es" ? "Personalice el nombre, eslogan y el estilo visual de toda la plataforma" : "Customize clinic name, slogan and primary color brand style"}
                      </p>
                    </div>
                  </div>

                  <div className="space-y-4">
                    <div>
                      <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5">
                        {lang === "es" ? "Nombre de la Clínica" : "Clinic Name"}
                      </label>
                      <input
                        type="text"
                        value={clinicName}
                        onChange={(e) => {
                          setClinicName(e.target.value);
                          localStorage.setItem("dianasrl_clinic_name", e.target.value);
                        }}
                        placeholder="Ej. DianaSRL"
                        className={inputStyle}
                      />
                    </div>

                    <div>
                      <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5">
                        {lang === "es" ? "Eslogan Corporativo" : "Clinic Slogan"}
                      </label>
                      <input
                        type="text"
                        value={clinicSlogan}
                        onChange={(e) => {
                          setClinicSlogan(e.target.value);
                          localStorage.setItem("dianasrl_clinic_slogan", e.target.value);
                        }}
                        placeholder="Ej. Odontología Estética de Vanguardia"
                        className={inputStyle}
                      />
                    </div>

                    {/* Brand Presets */}
                    <div>
                      <label className="block text-xs font-bold uppercase text-gray-500 mb-2.5">
                        {lang === "es" ? "Paleta de Color de la Marca" : "Brand Color Theme Preset"}
                      </label>
                      <div className="grid grid-cols-2 md:grid-cols-5 gap-2.5">
                        {[
                          { id: "pink", label: lang === "es" ? "Rosado Pastel" : "Pink Rose", colorClass: "bg-pink-500" },
                          { id: "emerald", label: lang === "es" ? "Esmeralda Dental" : "Dental Emerald", colorClass: "bg-emerald-500" },
                          { id: "amber", label: lang === "es" ? "Ámbar Cálido" : "Warm Amber", colorClass: "bg-amber-500" },
                          { id: "violet", label: lang === "es" ? "Violeta Amatista" : "Amethyst Violet", colorClass: "bg-violet-500" },
                          { id: "sky", label: lang === "es" ? "Azul Cielo" : "Sky Trust", colorClass: "bg-sky-500" },
                        ].map((preset) => {
                          const isPresetActive = clinicColorTheme === preset.id;
                          return (
                            <button
                              key={preset.id}
                              type="button"
                              onClick={() => {
                                setClinicColorTheme(preset.id as any);
                                localStorage.setItem("dianasrl_clinic_color_theme", preset.id);
                                triggerToast(lang === "es" ? `Tema cambiado a ${preset.label}` : `Theme changed to ${preset.label}`, "success");
                              }}
                              className={`p-3 rounded-2xl border text-center flex flex-col items-center justify-center space-y-2 transition-all ${
                                isPresetActive
                                  ? `${isDark ? "bg-white/5 border-pink-500" : "bg-rose-500/5 border-pink-500"} shadow-md scale-[1.03]`
                                  : "border-gray-200/50 dark:border-pink-500/5 hover:border-gray-300 dark:hover:border-pink-500/20"
                              }`}
                            >
                              <span className={`w-5 h-5 rounded-full ${preset.colorClass} shadow-sm block`} />
                              <span className="text-[10px] font-bold block truncate max-w-full text-gray-500 dark:text-pink-100/70">
                                {preset.label}
                              </span>
                            </button>
                          );
                        })}
                      </div>
                    </div>
                  </div>
                </div>

                {/* Feature Toggles Card (Quitar o Apagar cosas) */}
                <div className={bgCard}>
                  <div className="flex items-center space-x-2.5 border-b pb-4 mb-5 border-pink-500/10">
                    <Settings className={activeThemeObj.accentText} size={22} />
                    <div>
                      <h3 className={`text-lg font-bold ${textTitle}`}>
                        {lang === "es" ? "Gestión y Toggles de Funciones" : "Feature Toggles & Killswitches"}
                      </h3>
                      <p className="text-[10px] text-gray-400 font-medium">
                        {lang === "es" ? "Encienda o apague módulos completos del sistema de forma instantánea" : "Instantly enable or disable complete system features"}
                      </p>
                    </div>
                  </div>

                  <div className="space-y-3">
                    {[
                      { 
                        id: "aiDiagnostic", 
                        title: lang === "es" ? "Asistente Clínico Inteligente (IA)" : "AI Clinical Diagnostic Assistant",
                        desc: lang === "es" ? "Habilita la consulta interactiva de IA y el pre-diagnóstico de radiografías/dientes." : "Enables interactive AI consultation and radiography pre-diagnosis."
                      },
                      { 
                        id: "payments", 
                        title: lang === "es" ? "Módulo de Pagos y Caja" : "Cashier & Payments Ledger",
                        desc: lang === "es" ? "Habilita el registro de ingresos de caja, facturas, abonos y enlace de Stripe." : "Enables register of cash entries, invoices, deposits, and Stripe billing."
                      },
                      { 
                        id: "liveSync", 
                        title: lang === "es" ? "Sincronización en la Nube (Google Sheets)" : "Cloud Sync (Google Sheets)",
                        desc: lang === "es" ? "Habilita el centro de sincronización, importación y exportación de registros clínicos." : "Enables sync ledger, import/export clinical values to Sheets."
                      },
                      { 
                        id: "onlineRequests", 
                        title: lang === "es" ? "Formulario de Reserva Online" : "Online Appointment Booking Request",
                        desc: lang === "es" ? "Permite a los pacientes solicitar citas directamente desde su portal cliente." : "Allows patients to request dental appointments from their client dashboard."
                      }
                    ].map((feat) => {
                      const isFeatActive = activeFeatures[feat.id] !== false;
                      return (
                        <div 
                          key={feat.id}
                          className={`p-4 rounded-2xl border ${borderCol} flex items-start justify-between gap-4 transition-all ${
                            isFeatActive 
                              ? (isDark ? "bg-white/[0.02]" : "bg-pink-500/[0.01]") 
                              : "opacity-60 bg-gray-500/[0.02]"
                          }`}
                        >
                          <div className="space-y-1">
                            <span className="text-xs font-bold text-gray-800 dark:text-pink-100 block">
                              {feat.title}
                            </span>
                            <span className="text-[10px] text-gray-500 dark:text-pink-300/50 block leading-relaxed max-w-xl">
                              {feat.desc}
                            </span>
                          </div>

                          <button
                            type="button"
                            onClick={() => {
                              const updated = { ...activeFeatures, [feat.id]: !isFeatActive };
                              setActiveFeatures(updated);
                              localStorage.setItem("dianasrl_active_features", JSON.stringify(updated));
                              triggerToast(
                                lang === "es" 
                                  ? `${feat.title} ha sido ${!isFeatActive ? "ENCENDIDO" : "APAGADO"}` 
                                  : `${feat.title} has been ${!isFeatActive ? "ENABLED" : "DISABLED"}`,
                                !isFeatActive ? "success" : "error"
                              );
                            }}
                            className="shrink-0 transition-all duration-300 transform hover:scale-105"
                          >
                            {isFeatActive ? (
                              <ToggleRight className="text-pink-500" size={38} />
                            ) : (
                              <ToggleLeft className="text-gray-400" size={38} />
                            )}
                          </button>
                        </div>
                      );
                    })}
                  </div>
                </div>
              </div>

              {/* TERMS & CONDITIONS / POLICIES SECTION */}
              <div className="lg:col-span-2 space-y-6">
                <div className={bgCard}>
                  <div className="flex items-center space-x-2.5 border-b pb-4 mb-5 border-pink-500/10">
                    <ShieldCheck className={activeThemeObj.accentText} size={22} />
                    <div>
                      <h3 className={`text-lg font-bold ${textTitle}`}>
                        {lang === "es" ? "Políticas y Contratos" : "Terms & Policies Management"}
                      </h3>
                      <p className="text-[10px] text-gray-400 font-medium">
                        {lang === "es" ? "Actualice los términos legales expuestos a sus pacientes en tiempo real" : "Update legal terms, policies and contracts instantly for patients"}
                      </p>
                    </div>
                  </div>

                  <div className="space-y-4">
                    <div>
                      <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5 flex items-center space-x-1.5">
                        <span>{lang === "es" ? "Términos y Condiciones Legales" : "Terms & Conditions Contract"}</span>
                        <span className="text-[9px] bg-pink-500/10 text-pink-500 px-1.5 py-0.5 rounded font-mono font-bold uppercase">Paciente-Vista</span>
                      </label>
                      <textarea
                        rows={8}
                        value={termsAndConditions}
                        onChange={(e) => {
                          setTermsAndConditions(e.target.value);
                          localStorage.setItem("dianasrl_terms_conditions", e.target.value);
                        }}
                        placeholder="Escriba los términos legales..."
                        className={`${inputStyle} text-xs font-mono resize-none leading-relaxed`}
                      />
                    </div>

                    <div>
                      <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5 flex items-center space-x-1.5">
                        <span>{lang === "es" ? "Políticas de Privacidad de Datos" : "Privacy Policy Agreement"}</span>
                        <span className="text-[9px] bg-pink-500/10 text-pink-500 px-1.5 py-0.5 rounded font-mono font-bold uppercase">Paciente-Vista</span>
                      </label>
                      <textarea
                        rows={8}
                        value={privacyPolicy}
                        onChange={(e) => {
                          setPrivacyPolicy(e.target.value);
                          localStorage.setItem("dianasrl_privacy_policy", e.target.value);
                        }}
                        placeholder="Escriba las políticas de privacidad..."
                        className={`${inputStyle} text-xs font-mono resize-none leading-relaxed`}
                      />
                    </div>

                    <div className="pt-2">
                      <button
                        type="button"
                        onClick={() => {
                          triggerToast(
                            lang === "es" 
                              ? "¡Términos, Condiciones y Políticas guardadas con éxito!" 
                              : "Terms, conditions and privacy policies saved successfully!",
                            "success"
                          );
                        }}
                        className={`w-full py-3 bg-gradient-to-r ${activeThemeObj.gradient} text-white font-bold rounded-xl text-xs uppercase tracking-wider transition-all flex items-center justify-center space-x-2 shadow-md ${activeThemeObj.shadow}`}
                      >
                        <CheckCircle size={14} />
                        <span>{lang === "es" ? "Guardar y Publicar Cambios" : "Save and Publish Changes"}</span>
                      </button>
                    </div>
                  </div>
                </div>
              </div>
            </motion.div>
          )}

        </AnimatePresence>
        </>
        )
        )}

        {/* ========================================================= */}
        {/* PATIENT PORTAL (CLIENT VIEW)                              */}
        {/* ========================================================= */}
        {portalMode === "patient" && (
          <div className="space-y-8 animate-fade-in">
            
            {/* If NOT identified / logged in */}
            {!loggedPatientId ? (
              <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
                
                {/* Panel 1: Identify / Login */}
                <div className={bgCard}>
                  <div className="flex items-center space-x-2 border-b pb-4 mb-5 border-pink-500/10">
                    <Users className="text-pink-500" size={20} />
                    <h3 className={`text-lg font-bold ${textTitle}`}>
                      {lang === "es" ? "Acceso de Pacientes Registrados" : "Registered Patient Access"}
                    </h3>
                  </div>
                  
                  <p className="text-xs text-gray-500 mb-6 leading-relaxed">
                    {lang === "es" 
                      ? "Ingrese su número de cédula, DNI, pasaporte o su nombre completo para acceder a su expediente clínico, historial de odontograma, facturas de caja y consultas de Inteligencia Artificial."
                      : "Enter your DNI/Cédula, passport number or your full name to access your clinical file, odontogram history, cashier bills, and AI consultations."}
                  </p>

                  <div className="space-y-4">
                    <div>
                      <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5">
                        {lang === "es" ? "Cédula, DNI, Pasaporte o Nombre" : "Cédula, DNI, Passport or Name"}
                      </label>
                      <div className="relative">
                        <input
                          type="text"
                          placeholder={lang === "es" ? "Ej. 402-1234567-8 o Diana" : "e.g., 402-1234567-8 or Diana"}
                          value={patientDniInput}
                          onChange={e => setPatientDniInput(e.target.value)}
                          onKeyDown={e => {
                            if (e.key === "Enter") handlePatientLogin(patientDniInput);
                          }}
                          className={inputStyle}
                        />
                        <button
                          onClick={() => handlePatientLogin(patientDniInput)}
                          className="absolute right-2 top-1.5 p-2 bg-gradient-to-r from-pink-400 to-rose-500 hover:from-pink-500 hover:to-rose-600 text-white rounded-lg transition-all"
                        >
                          <Search size={14} />
                        </button>
                      </div>
                    </div>

                    <button
                      onClick={() => handlePatientLogin(patientDniInput)}
                      className="w-full py-3 bg-gradient-to-r from-pink-400 to-rose-500 hover:from-pink-500 hover:to-rose-600 text-white font-bold rounded-xl text-xs uppercase tracking-wider transition-all flex items-center justify-center space-x-2"
                    >
                      <Users size={14} />
                      <span>{lang === "es" ? "Ingresar al Portal" : "Enter Portal"}</span>
                    </button>

                    {/* Simulation accounts pills */}
                    <div className="space-y-2 mt-6 border-t border-pink-500/5 pt-4">
                      <p className="text-[10px] text-gray-400 uppercase font-bold tracking-wider">
                        {lang === "es" ? "Cuentas de Pacientes Existentes (Simulador):" : "Existing Patient Accounts (Simulator):"}
                      </p>
                      {patients.length === 0 ? (
                        <p className="text-xs text-gray-400 italic">No hay pacientes registrados en el sistema. ¡Regístrese a la derecha!</p>
                      ) : (
                        <div className="flex flex-wrap gap-2 max-h-40 overflow-y-auto pr-1">
                          {patients.map(p => (
                            <button
                              key={p.id}
                              onClick={() => {
                                setLoggedPatientId(p.id);
                                localStorage.setItem("dianasrl_logged_patient_id", p.id);
                                triggerToast(lang === "es" ? `¡Bienvenido, ${p.name}!` : `Welcome, ${p.name}!`, "success");
                              }}
                              className={`px-3 py-1.5 border ${borderCol} ${isDark ? "bg-pink-950/20 hover:bg-pink-950/40 text-pink-300" : "bg-rose-50/40 hover:bg-rose-100/60 text-rose-800"} text-xs rounded-xl transition-all font-semibold flex items-center space-x-1`}
                            >
                              <span className="w-1.5 h-1.5 bg-emerald-500 rounded-full" />
                              <span>{p.name}</span>
                            </button>
                          ))}
                        </div>
                      )}
                    </div>

                  </div>
                </div>

                {/* Panel 2: Register as New Patient */}
                <div className={bgCard}>
                  <div className="flex items-center space-x-2 border-b pb-4 mb-5 border-pink-500/10">
                    <PlusCircle className="text-pink-500" size={20} />
                    <h3 className={`text-lg font-bold ${textTitle}`}>
                      {lang === "es" ? "Ficha de Registro & Admisión Dental" : "Dental Admission & Intake Registration"}
                    </h3>
                  </div>

                  <form onSubmit={handlePatientRegister} className="space-y-6">
                    {/* Section 1: Personal info */}
                    <div className="space-y-4">
                      <h4 className="text-xs font-black uppercase text-pink-500 tracking-wider flex items-center space-x-1">
                        <span className="w-1.5 h-1.5 bg-pink-500 rounded-full" />
                        <span>{lang === "es" ? "1. Datos Personales y de Contacto" : "1. Personal & Contact Details"}</span>
                      </h4>

                      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                        <div className="md:col-span-2">
                          <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5">{lang === "es" ? "Nombre Completo" : "Full Name"} *</label>
                          <input
                            type="text"
                            required
                            placeholder="Ej. Camila Santos Ortiz"
                            value={patientRegForm.name}
                            onChange={e => setPatientRegForm({ ...patientRegForm, name: e.target.value })}
                            className={inputStyle}
                          />
                        </div>

                        <div>
                          <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5">{lang === "es" ? "Cédula / DNI / Pasaporte" : "DNI / National ID / Passport"} *</label>
                          <input
                            type="text"
                            required
                            placeholder={lang === "es" ? "Ej. 402-9876543-2 o RD123456" : "e.g., 402-9876543-2 or RD123456"}
                            value={patientRegForm.dni}
                            onChange={e => setPatientRegForm({ ...patientRegForm, dni: e.target.value })}
                            className={inputStyle}
                          />
                        </div>

                        <div>
                          <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5">{lang === "es" ? "Fecha de Nacimiento" : "Date of Birth"}</label>
                          <input
                            type="date"
                            value={patientRegForm.dob}
                            onChange={e => setPatientRegForm({ ...patientRegForm, dob: e.target.value })}
                            className={inputStyle}
                          />
                        </div>

                        <div>
                          <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5">{lang === "es" ? "Teléfono / WhatsApp" : "Phone / WhatsApp"}</label>
                          <input
                            type="text"
                            placeholder="Ej. 809-555-0199"
                            value={patientRegForm.phone}
                            onChange={e => setPatientRegForm({ ...patientRegForm, phone: e.target.value })}
                            className={inputStyle}
                          />
                        </div>

                        <div>
                          <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5">{lang === "es" ? "Vía de Contacto Preferida" : "Preferred Contact Method"}</label>
                          <select
                            value={patientRegForm.preferredContact}
                            onChange={e => setPatientRegForm({ ...patientRegForm, preferredContact: e.target.value })}
                            className={inputStyle}
                          >
                            <option value="WhatsApp">WhatsApp</option>
                            <option value="Llamada">{lang === "es" ? "Llamada Telefónica" : "Phone Call"}</option>
                            <option value="Email">Email / Correo</option>
                          </select>
                        </div>

                        <div className="md:col-span-2">
                          <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5">{lang === "es" ? "Correo Electrónico" : "Email"}</label>
                          <input
                            type="email"
                            placeholder="Ej. camila@ejemplo.com"
                            value={patientRegForm.email}
                            onChange={e => setPatientRegForm({ ...patientRegForm, email: e.target.value })}
                            className={inputStyle}
                          />
                        </div>

                        <div className="md:col-span-2">
                          <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5">{lang === "es" ? "Dirección / Ubicación (Opcional)" : "Address / Location (Optional)"}</label>
                          <input
                            type="text"
                            placeholder={lang === "es" ? "Ej. Ensanche Naco, Santo Domingo, RD" : "e.g., Ensanche Naco, Santo Domingo, DR"}
                            value={patientRegForm.address}
                            onChange={e => setPatientRegForm({ ...patientRegForm, address: e.target.value })}
                            className={inputStyle}
                          />
                        </div>
                      </div>
                    </div>

                    {/* Section 2: Clinical info */}
                    <div className="space-y-4 pt-4 border-t border-pink-500/10">
                      <h4 className="text-xs font-black uppercase text-pink-500 tracking-wider flex items-center space-x-1">
                        <span className="w-1.5 h-1.5 bg-pink-500 rounded-full" />
                        <span>{lang === "es" ? "2. Ficha Clínica e Historial Médico" : "2. Clinical & Medical History"}</span>
                      </h4>

                      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                        <div>
                          <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5">{lang === "es" ? "Seguro Dental / Médico" : "Dental Insurance"}</label>
                          <select
                            value={patientRegForm.dentalInsurance}
                            onChange={e => setPatientRegForm({ ...patientRegForm, dentalInsurance: e.target.value })}
                            className={inputStyle}
                          >
                            <option value="">-- {lang === "es" ? "Particular (Sin Seguro)" : "Private (No Insurance)"} --</option>
                            <option value="ARS Humano">ARS Humano</option>
                            <option value="ARS Palic / Mapfre">ARS Mapfre / Palic</option>
                            <option value="ARS Universal">ARS Universal</option>
                            <option value="ARS Senasa">ARS Senasa (Privado/Contributivo)</option>
                            <option value="Otro Seguro">{lang === "es" ? "Otro Seguro Médico" : "Other Insurance"}</option>
                          </select>
                        </div>

                        <div>
                          <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5">{lang === "es" ? "Motivo Principal de Consulta" : "Primary Consultation Reason"}</label>
                          <select
                            value={patientRegForm.consultationReason}
                            onChange={e => setPatientRegForm({ ...patientRegForm, consultationReason: e.target.value })}
                            className={inputStyle}
                          >
                            <option value="">-- {lang === "es" ? "Seleccione motivo" : "Select reason"} --</option>
                            <option value="Estética / Carillas">{lang === "es" ? "Estética Dental & Carillas" : "Cosmetic Dentistry & Veneers"}</option>
                            <option value="Blanqueamiento">{lang === "es" ? "Blanqueamiento Dental" : "Tooth Whitening"}</option>
                            <option value="Limpieza / Profilaxis">{lang === "es" ? "Limpieza & Profilaxis" : "Cleaning & Prophylaxis"}</option>
                            <option value="Dolor / Emergencia">{lang === "es" ? "Dolor o Emergencia" : "Toothache / Emergency"}</option>
                            <option value="Evaluación General">{lang === "es" ? "Evaluación General & Odontograma" : "General Checkup & Odontogram"}</option>
                          </select>
                        </div>

                        <div>
                          <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5">{lang === "es" ? "Alergias a Medicamentos" : "Drug Allergies"}</label>
                          <input
                            type="text"
                            placeholder={lang === "es" ? "Ej. Penicilina, Anestésicos (o escriba 'Ninguna')" : "e.g. Penicillin, Anesthetics (or write 'None')"}
                            value={patientRegForm.allergies}
                            onChange={e => setPatientRegForm({ ...patientRegForm, allergies: e.target.value })}
                            className={inputStyle}
                          />
                        </div>

                        <div>
                          <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5">{lang === "es" ? "Enfermedades o Condiciones Médicas" : "Medical Conditions"}</label>
                          <input
                            type="text"
                            placeholder={lang === "es" ? "Ej. Hipertensión, Diabetes, Embarazo, Ninguna" : "e.g. Hypertension, Diabetes, Pregnancy, None"}
                            value={patientRegForm.medicalConditions}
                            onChange={e => setPatientRegForm({ ...patientRegForm, medicalConditions: e.target.value })}
                            className={inputStyle}
                          />
                        </div>

                        <div className="md:col-span-2">
                          <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5">{lang === "es" ? "Observaciones Adicionales / Síntomas" : "Additional Observations / Symptoms"}</label>
                          <textarea
                            rows={2}
                            placeholder={lang === "es" ? "Describa brevemente si tiene alguna molestia actual, sensibilidad o expectativa dental..." : "Briefly describe any current discomfort, sensitivity or expectations..."}
                            value={patientRegForm.notes}
                            onChange={e => setPatientRegForm({ ...patientRegForm, notes: e.target.value })}
                            className={`${inputStyle} resize-none`}
                          />
                        </div>
                      </div>
                    </div>

                    <div className="pt-2">
                      <button
                        type="submit"
                        className="w-full py-3.5 bg-gradient-to-r from-pink-400 to-rose-500 hover:from-pink-500 hover:to-rose-600 text-white font-bold rounded-xl text-xs uppercase tracking-wider transition-all flex items-center justify-center space-x-2 shadow-md shadow-pink-500/10"
                      >
                        <Plus size={14} />
                        <span>{lang === "es" ? "Completar Registro e Ingresar" : "Complete Registration and Enter"}</span>
                      </button>
                    </div>
                  </form>
                </div>
              </div>
            ) : (
              
              /* Logged In Patient Experience */
              (() => {
                const loggedPatient = patients.find(p => p.id === loggedPatientId);
                if (!loggedPatient) return null;

                const patientAppointments = appointments.filter(a => a.patientId === loggedPatient.id);
                const nextAppt = patientAppointments.find(a => a.status === "Scheduled");
                
                const patientInvoices = payments.filter(p => 
                  p.patientId === loggedPatient.id || 
                  p.patientName.toLowerCase() === loggedPatient.name.toLowerCase()
                );
                const billingTotal = patientInvoices.reduce((sum, inv) => sum + inv.amountTotal, 0);
                const billingPaid = patientInvoices.reduce((sum, inv) => sum + inv.amountPaid, 0);
                const billingPending = patientInvoices.reduce((sum, inv) => sum + inv.amountPending, 0);

                return (
                  <div className="space-y-8">
                    
                    {/* Welcome Banner Card */}
                    <div className={`p-6 rounded-3xl border ${borderCol} ${isDark ? "bg-gradient-to-r from-pink-950/20 via-[#271a1c] to-[#1c1214]" : "bg-gradient-to-r from-rose-50/50 via-white to-rose-50/20"} flex flex-col md:flex-row items-center justify-between gap-6 relative overflow-hidden`}>
                      <div className="absolute top-0 right-0 w-32 h-32 bg-pink-400/10 dark:bg-pink-500/5 rounded-full blur-2xl pointer-events-none" />
                      
                      <div className="space-y-1.5 z-10">
                        <div className="inline-flex items-center space-x-2 px-3 py-1 rounded-full bg-pink-100 dark:bg-pink-950/30 text-pink-600 dark:text-pink-400 border border-pink-200/50 dark:border-pink-500/10">
                          <Heart size={10} className="fill-current text-rose-500" />
                          <span className="text-[10px] font-bold uppercase tracking-widest font-mono">{clinicName} {lang === "es" ? "Portal de Pacientes" : "Patient Portal"}</span>
                        </div>
                        <h2 className={`text-2xl md:text-3xl font-extrabold tracking-tight ${textTitle}`}>
                          {lang === "es" ? `¡Hola, ${loggedPatient.name}!` : `Hello, ${loggedPatient.name}!`}
                        </h2>
                        <p className={`text-xs ${textSub} max-w-xl leading-relaxed`}>
                          {lang === "es" 
                            ? "Bienvenido a su área personal. Desde aquí puede solicitar citas estéticas, ver las observaciones de la Dra. Diana Rojas en su mapa de salud dental, descargar facturas de pago y consultar con nuestro asistente de IA."
                            : "Welcome to your personal dashboard. From here you can request appointments, check dentist remarks on your dental map, download invoices, and consult with our dental AI advisor."}
                        </p>
                      </div>

                      <div className="flex flex-col sm:flex-row items-center gap-4 z-10 shrink-0">
                        {/* Quick metrics */}
                        <div className={`px-4 py-3 rounded-2xl border ${borderCol} ${isDark ? "bg-pink-950/30" : "bg-white"} text-center`}>
                          <div className="text-xs font-extrabold uppercase tracking-widest text-pink-500 font-mono">
                            {lang === "es" ? "Estado de Cuenta" : "Account Balance"}
                          </div>
                          <div className="text-lg font-bold text-gray-950 dark:text-pink-100 mt-1">
                            {billingPending > 0 ? (
                              <span className="text-rose-500 font-extrabold">${billingPending}</span>
                            ) : (
                              <span className="text-emerald-500 font-extrabold">Al día</span>
                            )}
                          </div>
                        </div>

                        <button
                          onClick={handlePatientLogout}
                          className="px-4 py-3 bg-rose-500 hover:bg-rose-600 text-white rounded-2xl font-bold text-xs uppercase tracking-wider transition-all flex items-center justify-center space-x-2 shadow-lg shadow-rose-200 dark:shadow-none"
                          title="Cerrar sesión"
                        >
                          <LogOut size={14} />
                          <span>{lang === "es" ? "Salir" : "Logout"}</span>
                        </button>
                      </div>
                    </div>

                    {/* Sub-Views Content Tabs */}
                    <AnimatePresence mode="wait">
                      
                      {/* TAB: Patient Appointments */}
                      {patientActiveTab === "my-appointments" && (
                        <motion.div
                          key="patient-appointments"
                          initial={{ opacity: 0, y: 15 }}
                          animate={{ opacity: 1, y: 0 }}
                          exit={{ opacity: 0, y: -15 }}
                          className="grid grid-cols-1 lg:grid-cols-3 gap-8"
                        >
                          {/* Request Appointment Form */}
                          <div className={`${bgCard} lg:col-span-1 h-fit`}>
                            <div className="flex items-center space-x-2 border-b pb-4 mb-5 border-pink-500/10">
                              <Calendar className="text-pink-500" size={20} />
                              <h3 className={`text-lg font-bold ${textTitle}`}>
                                {lang === "es" ? "Solicitar Cita Estética" : "Request Dental Appointment"}
                              </h3>
                            </div>

                            {activeFeatures.onlineRequests === false ? (
                              <div className="p-5 bg-amber-500/5 dark:bg-amber-950/15 border border-amber-500/10 rounded-2xl space-y-3">
                                <div className="w-10 h-10 bg-amber-500/10 rounded-xl flex items-center justify-center">
                                  <AlertTriangle className="text-amber-500" size={20} />
                                </div>
                                <h4 className="text-sm font-bold text-amber-500 uppercase tracking-wide">
                                  {lang === "es" ? "Solicitudes Inactivas" : "Scheduling Offline"}
                                </h4>
                                <p className="text-xs text-gray-500 dark:text-pink-300/60 leading-relaxed">
                                  {lang === "es" 
                                    ? "Las solicitudes de citas online se encuentran temporalmente inactivas. Por favor, contáctenos directamente vía WhatsApp o llamada telefónica para reservar su espacio de inmediato."
                                    : "Online appointment requests are temporarily disabled. Please contact us directly via WhatsApp or phone call to book your slot immediately."}
                                </p>
                              </div>
                            ) : (
                              <form onSubmit={handlePatientRequestAppt} className="space-y-4">
                              <div>
                                <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5">Especialista Requerido</label>
                                <select
                                  value={patientNewAppt.dentistName}
                                  onChange={e => setPatientNewAppt({ ...patientNewAppt, dentistName: e.target.value })}
                                  className={inputStyle}
                                >
                                  <option value="Dra. Diana Rojas">Dra. Diana Rojas (Estética Dental)</option>
                                  <option value="Dr. Marcos Soler">Dr. Marcos Soler (Cirugía & Implantes)</option>
                                  <option value="Dra. Lucía Santos">Dra. Lucía Santos (Endodoncia)</option>
                                </select>
                              </div>

                              <div>
                                <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5">Tratamiento de Interés *</label>
                                <select
                                  required
                                  value={patientNewAppt.treatmentType}
                                  onChange={e => setPatientNewAppt({ ...patientNewAppt, treatmentType: e.target.value })}
                                  className={inputStyle}
                                >
                                  <option value="">Seleccione un tratamiento...</option>
                                  <option value="Carillas de Disilicato">Carillas de Disilicato de Litio</option>
                                  <option value="Blanqueamiento Láser">Blanqueamiento Láser Estético</option>
                                  <option value="Resina Estética">Resina Estética de Alta Estética</option>
                                  <option value="Limpieza Profiláctica">Limpieza Profiláctica Completa</option>
                                  <option value="Corona de Zirconio">Corona Estética de Zirconio</option>
                                  <option value="Implante Titanio">Implante Dental de Titanio</option>
                                </select>
                              </div>

                              {/* Integrated Availability System */}
                              {(() => {
                                const docName = patientNewAppt.dentistName || "Dra. Diana Rojas";
                                const docAvails = availabilities.filter(
                                  a => a.doctorName.toLowerCase() === docName.toLowerCase() && a.slots.length > 0
                                );

                                return (
                                  <div className="space-y-4">
                                    <div>
                                      <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5">
                                        {lang === "es" ? "Fecha de la Consulta *" : "Consultation Date *"}
                                      </label>
                                      {docAvails.length > 0 ? (
                                        <select
                                          required
                                          value={patientNewAppt.date}
                                          onChange={e => setPatientNewAppt({ ...patientNewAppt, date: e.target.value, time: "" })}
                                          className={inputStyle}
                                        >
                                          <option value="">-- {lang === "es" ? "Seleccione una fecha programada" : "Select a scheduled date"} --</option>
                                          {docAvails
                                            .sort((a, b) => a.date.localeCompare(b.date))
                                            .map(avail => (
                                              <option key={avail.id} value={avail.date}>
                                                {avail.date} ({lang === "es" ? "Atención Disponible" : "Care Available"})
                                              </option>
                                            ))}
                                        </select>
                                      ) : (
                                        <div>
                                          <input
                                            type="date"
                                            required
                                            min={new Date().toISOString().split("T")[0]}
                                            value={patientNewAppt.date}
                                            onChange={e => setPatientNewAppt({ ...patientNewAppt, date: e.target.value, time: "" })}
                                            className={inputStyle}
                                          />
                                          <p className="text-[10px] text-pink-500 font-semibold mt-1">
                                            ⚠️ {lang === "es" ? "Horario estándar activo (Doctor sin agenda personalizada para hoy)" : "Standard schedule active (Doctor has no customized agenda today)"}
                                          </p>
                                        </div>
                                      )}
                                    </div>

                                    {patientNewAppt.date && (
                                      <div>
                                        <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5">
                                          {lang === "es" ? "Horarios Disponibles *" : "Available Times *"}
                                        </label>
                                        {(() => {
                                          const matchingAvail = availabilities.find(
                                            a => a.doctorName.toLowerCase() === docName.toLowerCase() && a.date === patientNewAppt.date
                                          );

                                          // Get already booked times for this doctor on this date
                                          const bookedTimes = appointments
                                            .filter(
                                              app =>
                                                app.dentistName.toLowerCase() === docName.toLowerCase() &&
                                                app.date === patientNewAppt.date &&
                                                app.status === "Scheduled"
                                            )
                                            .map(app => app.time);

                                          const availableSlots = matchingAvail
                                            ? matchingAvail.slots.filter(s => !bookedTimes.includes(s))
                                            : ["08:00", "09:00", "10:00", "11:00", "12:00", "14:00", "15:00", "16:00", "17:00"].filter(s => !bookedTimes.includes(s));

                                          if (availableSlots.length === 0) {
                                            return (
                                              <p className="text-xs text-rose-500 font-bold bg-rose-500/10 p-3 rounded-xl border border-rose-500/10 text-center">
                                                ❌ {lang === "es" ? "No hay horarios libres para esta fecha." : "No free slots left for this date."}
                                              </p>
                                            );
                                          }

                                          return (
                                            <div className="grid grid-cols-3 gap-2 p-3 bg-pink-500/5 dark:bg-black/20 rounded-2xl border border-pink-500/10">
                                              {availableSlots.map(time => {
                                                const isSelected = patientNewAppt.time === time;
                                                return (
                                                  <button
                                                    key={time}
                                                    type="button"
                                                    onClick={() => setPatientNewAppt({ ...patientNewAppt, time })}
                                                    className={`py-1.5 text-xs font-mono font-bold rounded-lg transition-all cursor-pointer ${
                                                      isSelected
                                                        ? "bg-gradient-to-r from-pink-400 to-rose-500 text-white shadow shadow-pink-500/10 scale-[1.03]"
                                                        : "bg-white dark:bg-[#1f1618] text-rose-800 dark:text-pink-300 border border-pink-500/10 hover:border-pink-500/40"
                                                    }`}
                                                  >
                                                    {time}
                                                  </button>
                                                );
                                              })}
                                            </div>
                                          );
                                        })()}
                                      </div>
                                    )}
                                  </div>
                                );
                              })()}

                              <div>
                                <label className="block text-xs font-bold uppercase text-gray-500 mb-1.5">Notas o Detalles de Preferencia</label>
                                <textarea
                                  rows={3}
                                  placeholder="Ej. Me gustaría programar en la mañana o detallar molestias anteriores."
                                  value={patientNewAppt.notes}
                                  onChange={e => setPatientNewAppt({ ...patientNewAppt, notes: e.target.value })}
                                  className={`${inputStyle} resize-none`}
                                />
                              </div>

                              <button
                                type="submit"
                                className="w-full py-3 bg-gradient-to-r from-pink-400 to-rose-500 hover:from-pink-500 hover:to-rose-600 text-white font-bold rounded-xl text-xs uppercase tracking-wider transition-all flex items-center justify-center space-x-2 shadow-md shadow-pink-500/10"
                              >
                                <Calendar size={14} />
                                <span>Agendar mi Cita</span>
                              </button>
                            </form>
                            )}
                          </div>

                          {/* Dental Intake Profile Summary */}
                          <div className={`${bgCard} lg:col-span-1 h-fit mt-6`}>
                            <div className="flex items-center space-x-2 border-b pb-4 mb-4 border-pink-500/10">
                              <Heart className="text-pink-500 fill-pink-500/10" size={18} />
                              <h3 className={`text-sm font-bold ${textTitle}`}>
                                {lang === "es" ? "Mi Ficha de Admisión Dental" : "My Clinical Intake File"}
                              </h3>
                            </div>

                            <div className="space-y-4">
                              <div className="grid grid-cols-2 gap-3 text-xs font-semibold text-gray-500 dark:text-pink-300/60">
                                <div className="col-span-2 md:col-span-1">
                                  <span className="text-[10px] text-gray-400 block uppercase font-bold">{lang === "es" ? "Identificación (DNI/Pasaporte)" : "DNI / Passport ID"}</span>
                                  <span className={isDark ? "text-pink-100" : "text-rose-950 font-mono"}>{loggedPatient.dni}</span>
                                </div>
                                <div className="col-span-2 md:col-span-1">
                                  <span className="text-[10px] text-gray-400 block uppercase font-bold">{lang === "es" ? "F. Nacimiento" : "Birth Date"}</span>
                                  <span className={isDark ? "text-pink-100" : "text-rose-950"}>{loggedPatient.dob || "--"}</span>
                                </div>
                                <div>
                                  <span className="text-[10px] text-gray-400 block uppercase font-bold">{lang === "es" ? "Contacto" : "Contact Channel"}</span>
                                  <span className="text-pink-500 font-bold">{loggedPatient.preferredContact || "WhatsApp"}</span>
                                </div>
                                <div>
                                  <span className="text-[10px] text-gray-400 block uppercase font-bold">{lang === "es" ? "Seguro Dental" : "Insurance"}</span>
                                  <span className={isDark ? "text-pink-100" : "text-rose-950"}>{loggedPatient.dentalInsurance || (lang === "es" ? "Particular" : "None")}</span>
                                </div>
                              </div>

                              {loggedPatient.address && (
                                <div className="text-xs pt-2.5 border-t border-pink-500/5">
                                  <span className="text-[10px] text-gray-400 block uppercase font-bold mb-0.5">{lang === "es" ? "Dirección / Ubicación" : "Address / Location"}</span>
                                  <p className={isDark ? "text-pink-200" : "text-rose-950"}>{loggedPatient.address}</p>
                                </div>
                              )}

                              {loggedPatient.consultationReason && (
                                <div className="text-xs pt-2.5 border-t border-pink-500/5">
                                  <span className="text-[10px] text-gray-400 block uppercase font-bold mb-0.5">{lang === "es" ? "Motivo Principal de Consulta" : "Consultation Reason"}</span>
                                  <p className={isDark ? "text-pink-200" : "text-rose-900"}>{loggedPatient.consultationReason}</p>
                                </div>
                              )}

                              {(loggedPatient.allergies || loggedPatient.medicalConditions) && (
                                <div className="p-3 rounded-2xl bg-amber-500/5 border border-amber-500/10 space-y-2 text-xs">
                                  <span className="text-[9px] font-black uppercase text-amber-600 dark:text-amber-400 tracking-wider block">
                                    ⚠️ {lang === "es" ? "Alertas Clínicas Registradas" : "Clinical Alerts On File"}
                                  </span>
                                  {loggedPatient.allergies && (
                                    <div>
                                      <span className="text-[10px] font-bold text-amber-700 dark:text-amber-400">{lang === "es" ? "Alergias a Medicamentos:" : "Drug Allergies:"}</span>
                                      <p className="text-gray-600 dark:text-pink-100/80 mt-0.5 font-medium italic">{loggedPatient.allergies}</p>
                                    </div>
                                  )}
                                  {loggedPatient.medicalConditions && (
                                    <div>
                                      <span className="text-[10px] font-bold text-amber-700 dark:text-amber-400">{lang === "es" ? "Condiciones de Salud:" : "Medical Conditions:"}</span>
                                      <p className="text-gray-600 dark:text-pink-100/80 mt-0.5 font-medium italic">{loggedPatient.medicalConditions}</p>
                                    </div>
                                  )}
                                </div>
                              )}

                              {loggedPatient.notes && (
                                <div className="text-xs pt-2.5 border-t border-pink-500/5">
                                  <span className="text-[10px] text-gray-400 block uppercase font-bold mb-0.5">{lang === "es" ? "Observaciones de Admisión" : "Admission Notes"}</span>
                                  <p className="text-gray-600 dark:text-pink-200/80 leading-relaxed italic bg-rose-50/10 dark:bg-black/10 p-2 rounded-xl border border-pink-500/5">
                                    "{loggedPatient.notes}"
                                  </p>
                                </div>
                              )}
                            </div>
                          </div>

                          {/* My Appointment List */}
                          <div className={`${bgCard} lg:col-span-2 flex flex-col`}>
                            <div className="flex items-center space-x-2 border-b pb-4 mb-4 border-pink-500/10">
                              <Clock className="text-pink-500" size={20} />
                              <h3 className={`text-lg font-bold ${textTitle}`}>
                                {lang === "es" ? "Mi Historial de Citas Médicas" : "My Appointment History"}
                              </h3>
                            </div>

                            {patientAppointments.length === 0 ? (
                              <div className="flex-grow flex flex-col items-center justify-center py-12 text-center space-y-4">
                                <div className="w-16 h-16 bg-rose-50 dark:bg-pink-950/20 rounded-full flex items-center justify-center text-rose-300 dark:text-pink-500">
                                  <Calendar size={32} />
                                </div>
                                <div className="space-y-1">
                                  <p className={`font-bold ${textTitle}`}>No hay citas programadas</p>
                                  <p className="text-xs text-gray-400 max-w-sm">No registra citas estéticas activas en este momento. Solicite su primera cita usando el formulario de la izquierda.</p>
                                </div>
                              </div>
                            ) : (
                              <div className="space-y-3 overflow-y-auto max-h-[500px] pr-2 custom-scrollbar">
                                {patientAppointments.map(appt => {
                                  const statusStyle = appt.status === "Completed"
                                    ? "bg-emerald-500/10 text-emerald-500 border-emerald-500/20"
                                    : appt.status === "Cancelled"
                                    ? "bg-red-500/10 text-red-500 border-red-500/20"
                                    : "bg-pink-500/10 text-pink-500 border-pink-500/20 animate-pulse";

                                  return (
                                    <div key={appt.id} className={`p-4 rounded-2xl border ${borderCol} ${isDark ? "bg-[#1f1618]" : "bg-rose-50/10"} flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4`}>
                                      <div className="space-y-1.5">
                                        <div className="flex items-center space-x-2">
                                          <span className={`px-2.5 py-0.5 rounded-full text-[9px] font-extrabold border uppercase tracking-wider ${statusStyle}`}>
                                            {appt.status === "Scheduled" ? "Programada" : appt.status === "Completed" ? "Completada" : "Cancelada"}
                                          </span>
                                          <span className="text-[10px] text-pink-500 font-bold font-mono uppercase tracking-wider">{appt.id.toUpperCase()}</span>
                                        </div>
                                        <h4 className={`text-sm font-bold ${textTitle}`}>{appt.treatmentType}</h4>
                                        <p className="text-xs text-gray-500 font-semibold flex items-center space-x-1">
                                          <Users size={12} className="text-pink-400 shrink-0" />
                                          <span>Especialista: {appt.dentistName}</span>
                                        </p>
                                        {appt.notes && (
                                          <p className="text-xs text-gray-400 bg-rose-100/10 dark:bg-black/10 p-2 rounded-lg italic border border-pink-500/5 mt-1">
                                            "{appt.notes}"
                                          </p>
                                        )}
                                      </div>

                                      <div className="text-left sm:text-right shrink-0">
                                        <p className={`font-mono text-xs font-bold ${textTitle}`}>{appt.date}</p>
                                        <p className="text-[10px] text-pink-500 font-mono font-bold">{appt.time}</p>
                                      </div>
                                    </div>
                                  );
                                })}
                              </div>
                            )}
                          </div>
                        </motion.div>
                      )}

                      {/* TAB: Patient Odontogram */}
                      {patientActiveTab === "my-odontogram" && (
                        <motion.div
                          key="patient-odontogram"
                          initial={{ opacity: 0, y: 15 }}
                          animate={{ opacity: 1, y: 0 }}
                          exit={{ opacity: 0, y: -15 }}
                          className="space-y-8"
                        >
                          {/* Odontogram interactive canvas card */}
                          <div className={bgCard}>
                            <div className="flex items-center space-x-2 border-b pb-4 mb-6 border-pink-500/10">
                              <Activity className="text-pink-500" size={20} />
                              <div>
                                <h3 className={`text-lg font-bold ${textTitle}`}>
                                  {lang === "es" ? "Mi Mapa de Salud Dental Interactivo" : "My Interactive Dental Map"}
                                </h3>
                                <p className="text-xs text-gray-500 mt-0.5">Haga clic en cualquiera de sus dientes para ver observaciones de su especialista DianaSRL y consejos de salud.</p>
                              </div>
                            </div>

                            {/* Main Interactive Teeth Arch Container */}
                            <div className="overflow-x-auto pb-4">
                              <div className="min-w-[700px] space-y-6">
                                
                                {/* Top teeth row (1 - 16) */}
                                <div className="space-y-2">
                                  <div className="text-center text-[10px] font-extrabold tracking-widest text-pink-500 uppercase font-mono">Arcada Superior (Upper Arch)</div>
                                  <div className="grid grid-cols-16 gap-2 py-4 bg-rose-50/10 dark:bg-pink-950/10 rounded-2xl border border-pink-500/5 px-4 justify-items-center">
                                    {Array.from({ length: 16 }, (_, i) => {
                                      const toothNum = i + 1;
                                      const conditions = toothConditions[loggedPatient.id] || [];
                                      const cond = conditions.find(c => c.toothNumber === toothNum);
                                      const status = cond?.status || "Healthy";
                                      const isSelected = patientSelectedTooth === toothNum;

                                      // Colors representing dental states beautifully
                                      const statusColors: Record<string, string> = {
                                        Healthy: "bg-emerald-400 hover:bg-emerald-500 text-white shadow-emerald-400/20",
                                        Caries: "bg-red-500 hover:bg-red-600 text-white shadow-red-500/20",
                                        Crown: "bg-amber-500 hover:bg-amber-600 text-white shadow-amber-500/20",
                                        Endo: "bg-purple-500 hover:bg-purple-600 text-white shadow-purple-500/20",
                                        Filling: "bg-sky-500 hover:bg-sky-600 text-white shadow-sky-500/20",
                                        Missing: "bg-gray-400 hover:bg-gray-500 text-white shadow-gray-400/20"
                                      };

                                      return (
                                        <button
                                          key={toothNum}
                                          onClick={() => setPatientSelectedTooth(toothNum)}
                                          className={`w-10 h-14 rounded-xl flex flex-col items-center justify-between py-1.5 transition-all transform hover:scale-110 font-mono text-xs font-bold shadow-md cursor-pointer ${statusColors[status]} ${
                                            isSelected ? "ring-4 ring-pink-500 ring-offset-2 dark:ring-offset-[#251b1d] scale-105" : ""
                                          }`}
                                        >
                                          <span className="text-[9px] opacity-75 font-bold">#{toothNum}</span>
                                          <div className="w-4 h-4 rounded-full bg-white/20 flex items-center justify-center text-[9px] uppercase font-bold">
                                            {status.slice(0, 2)}
                                          </div>
                                        </button>
                                      );
                                    })}
                                  </div>
                                </div>

                                {/* Bottom teeth row (17 - 32) */}
                                <div className="space-y-2">
                                  <div className="text-center text-[10px] font-extrabold tracking-widest text-pink-500 uppercase font-mono">Arcada Inferior (Lower Arch)</div>
                                  <div className="grid grid-cols-16 gap-2 py-4 bg-rose-50/10 dark:bg-pink-950/10 rounded-2xl border border-pink-500/5 px-4 justify-items-center">
                                    {Array.from({ length: 16 }, (_, i) => {
                                      const toothNum = 17 + i;
                                      const conditions = toothConditions[loggedPatient.id] || [];
                                      const cond = conditions.find(c => c.toothNumber === toothNum);
                                      const status = cond?.status || "Healthy";
                                      const isSelected = patientSelectedTooth === toothNum;

                                      const statusColors: Record<string, string> = {
                                        Healthy: "bg-emerald-400 hover:bg-emerald-500 text-white shadow-emerald-400/20",
                                        Caries: "bg-red-500 hover:bg-red-600 text-white shadow-red-500/20",
                                        Crown: "bg-amber-500 hover:bg-amber-600 text-white shadow-amber-500/20",
                                        Endo: "bg-purple-500 hover:bg-purple-600 text-white shadow-purple-500/20",
                                        Filling: "bg-sky-500 hover:bg-sky-600 text-white shadow-sky-500/20",
                                        Missing: "bg-gray-400 hover:bg-gray-500 text-white shadow-gray-400/20"
                                      };

                                      return (
                                        <button
                                          key={toothNum}
                                          onClick={() => setPatientSelectedTooth(toothNum)}
                                          className={`w-10 h-14 rounded-xl flex flex-col items-center justify-between py-1.5 transition-all transform hover:scale-110 font-mono text-xs font-bold shadow-md cursor-pointer ${statusColors[status]} ${
                                            isSelected ? "ring-4 ring-pink-500 ring-offset-2 dark:ring-offset-[#251b1d] scale-105" : ""
                                          }`}
                                        >
                                          <span className="text-[9px] opacity-75 font-bold">#{toothNum}</span>
                                          <div className="w-4 h-4 rounded-full bg-white/20 flex items-center justify-center text-[9px] uppercase font-bold">
                                            {status.slice(0, 2)}
                                          </div>
                                        </button>
                                      );
                                    })}
                                  </div>
                                </div>

                              </div>
                            </div>

                            {/* Map legends */}
                            <div className="mt-4 flex flex-wrap gap-4 text-xs font-bold uppercase justify-center border-t border-pink-500/5 pt-4 text-gray-500">
                              <span className="flex items-center space-x-1.5"><span className="w-3 h-3 bg-emerald-400 rounded-full" /> <span>Sano (Healthy)</span></span>
                              <span className="flex items-center space-x-1.5"><span className="w-3 h-3 bg-red-500 rounded-full" /> <span>Caries (Caries)</span></span>
                              <span className="flex items-center space-x-1.5"><span className="w-3 h-3 bg-amber-500 rounded-full" /> <span>Corona (Crown)</span></span>
                              <span className="flex items-center space-x-1.5"><span className="w-3 h-3 bg-purple-500 rounded-full" /> <span>Endodoncia (Endo)</span></span>
                              <span className="flex items-center space-x-1.5"><span className="w-3 h-3 bg-sky-500 rounded-full" /> <span>Resina (Filling)</span></span>
                              <span className="flex items-center space-x-1.5"><span className="w-3 h-3 bg-gray-400 rounded-full" /> <span>Ausente (Missing)</span></span>
                            </div>

                          </div>

                          {/* Selected Tooth Drawer Panel */}
                          <AnimatePresence mode="wait">
                            {patientSelectedTooth ? (
                              (() => {
                                const toothNum = patientSelectedTooth;
                                const conditions = toothConditions[loggedPatient.id] || [];
                                const cond = conditions.find(c => c.toothNumber === toothNum);
                                const status = cond?.status || "Healthy";
                                const notes = cond?.notes || "";

                                const adviceObj: Record<string, { label: string; tips: string[] }> = {
                                  Healthy: {
                                    label: "Pieza Completamente Sana",
                                    tips: [
                                      "Continúe cepillándose tres veces al día con pasta dental fluorada.",
                                      "Utilice hilo dental todas las noches antes de dormir.",
                                      "Asista a su profilaxis dental semestral con la Dra. Diana Rojas para pulido preventivo."
                                    ]
                                  },
                                  Caries: {
                                    label: "Pieza con Caries Registrada",
                                    tips: [
                                      "Se recomienda programar una cita para restauración con resina estética de alta calidad.",
                                      "Evite el consumo excesivo de azúcares refinados y bebidas gaseosas para evitar que progrese.",
                                      "Cepille cuidadosamente la zona afectada sin aplicar fuerza abrasiva extrema."
                                    ]
                                  },
                                  Crown: {
                                    label: "Corona Estética Registrada",
                                    tips: [
                                      "Evite masticar elementos sumamente duros como cubos de hielo o nueces directamente con este diente.",
                                      "Preste especial atención al hilo dental en el margen gingival de la corona para evitar acumulación de placa.",
                                      "Las coronas de Zirconio y Disilicato de DianaSRL tienen alta durabilidad; asista a chequeo anual."
                                    ]
                                  },
                                  Endo: {
                                    label: "Tratamiento de Conducto (Endodoncia)",
                                    tips: [
                                      "Este diente ya no posee nervio sensible, pero se debe monitorear radiográficamente de forma anual.",
                                      "Asegúrese de que el sellado de la pieza esté óptimo para evitar re-infecciones bacterianas.",
                                      "Idealmente, las piezas con endodoncia se benefician de una corona protectora para evitar fracturas."
                                    ]
                                  },
                                  Filling: {
                                    label: "Restauración con Resina Estética",
                                    tips: [
                                      "Las resinas utilizadas por la Dra. Diana Rojas son de altísima estética y mimetizan el esmalte dental.",
                                      "Para evitar tinciones, disminuya el consumo excesivo de café, vino tinto o alimentos con colorantes.",
                                      "El pulido regular durante sus limpiezas mantendrá el brillo natural de la resina."
                                    ]
                                  },
                                  Missing: {
                                    label: "Pieza Dental Ausente",
                                    tips: [
                                      "La falta de un diente provoca la migración de los dientes vecinos, afectando la oclusión y masticación.",
                                      "Le sugerimos agendar una consulta de evaluación para un implante de titanio estética premium.",
                                      "Los implantes oseointegrados detienen la pérdida ósea y restablecen su sonrisa de forma permanente."
                                    ]
                                  }
                                };

                                const currentAdvice = adviceObj[status];

                                return (
                                  <motion.div
                                    key={`tooth-info-${toothNum}`}
                                    initial={{ opacity: 0, scale: 0.95 }}
                                    animate={{ opacity: 1, scale: 1 }}
                                    exit={{ opacity: 0, scale: 0.95 }}
                                    className={bgCard}
                                  >
                                    <div className="flex flex-col md:flex-row justify-between gap-6">
                                      <div className="space-y-4 flex-grow">
                                        <div className="flex items-center space-x-3">
                                          <div className="w-12 h-12 rounded-2xl bg-pink-500/10 text-pink-500 flex items-center justify-center font-mono font-black text-lg">
                                            #{toothNum}
                                          </div>
                                          <div>
                                            <h4 className={`text-base font-bold ${textTitle}`}>Ficha de la Pieza Dental</h4>
                                            <div className="flex items-center space-x-2 mt-1">
                                              <span className="w-2.5 h-2.5 rounded-full bg-pink-500 animate-pulse" />
                                              <span className="text-xs font-extrabold text-pink-500 uppercase tracking-widest">{currentAdvice.label}</span>
                                            </div>
                                          </div>
                                        </div>

                                        <div className="space-y-1.5">
                                          <span className="text-[10px] font-bold text-gray-400 uppercase tracking-wider">Notas Clínicas del Especialista</span>
                                          <p className="text-sm text-gray-700 dark:text-pink-100 font-medium bg-rose-50/10 dark:bg-[#1f1618] p-3 rounded-xl border border-pink-500/5 italic leading-relaxed">
                                            {notes ? `"${notes}"` : '"La Dra. Diana Rojas no ha registrado notas de anomalías en este diente. Se reporta con anatomía y salud impecables."'}
                                          </p>
                                        </div>
                                      </div>

                                      <div className="md:w-1/2 p-4 rounded-2xl border border-pink-500/5 bg-rose-50/20 dark:bg-[#1a1113] space-y-3">
                                        <h5 className="text-xs font-bold uppercase tracking-wider text-pink-500 flex items-center space-x-1.5">
                                          <Info size={14} />
                                          <span>Consejos Especializados de Higiene</span>
                                        </h5>
                                        <ul className="space-y-2">
                                          {currentAdvice.tips.map((tip, idx) => (
                                            <li key={idx} className="text-xs text-gray-500 dark:text-pink-300/60 leading-relaxed flex items-start space-x-2">
                                              <span className="text-pink-500 mt-1">✦</span>
                                              <span>{tip}</span>
                                            </li>
                                          ))}
                                        </ul>
                                      </div>
                                    </div>
                                  </motion.div>
                                );
                              })()
                            ) : (
                              <div className={`p-6 text-center border border-dashed ${borderCol} rounded-3xl text-gray-400 text-xs font-bold uppercase tracking-wider`}>
                                Seleccione una pieza dental arriba para ver su historial clínico de salud y cuidado personalizado.
                              </div>
                            )}
                          </AnimatePresence>

                        </motion.div>
                      )}

                      {/* TAB: Patient Billing */}
                      {patientActiveTab === "my-billing" && (
                        <motion.div
                          key="patient-billing"
                          initial={{ opacity: 0, y: 15 }}
                          animate={{ opacity: 1, y: 0 }}
                          exit={{ opacity: 0, y: -15 }}
                          className="space-y-8"
                        >
                          {/* Financial totals banner row */}
                          <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                            <div className={`${bgCard} flex items-center justify-between`}>
                              <div>
                                <p className="text-xs font-bold uppercase text-gray-400 tracking-wider">Total Contratado</p>
                                <h4 className={`text-2xl font-black ${textTitle} mt-1`}>${billingTotal}</h4>
                              </div>
                              <div className="w-10 h-10 bg-pink-100 dark:bg-pink-950/40 rounded-xl flex items-center justify-center text-pink-500"><DollarSign size={20} /></div>
                            </div>

                            <div className={`${bgCard} flex items-center justify-between`}>
                              <div>
                                <p className="text-xs font-bold uppercase text-gray-400 tracking-wider">Monto Abonado</p>
                                <h4 className="text-2xl font-black text-emerald-500 mt-1">${billingPaid}</h4>
                              </div>
                              <div className="w-10 h-10 bg-emerald-50 dark:bg-emerald-950/20 rounded-xl flex items-center justify-center text-emerald-500"><CheckCircle size={20} /></div>
                            </div>

                            <div className={`${bgCard} flex items-center justify-between`}>
                              <div>
                                <p className="text-xs font-bold uppercase text-gray-400 tracking-wider">Balance Pendiente</p>
                                <h4 className={`text-2xl font-black ${billingPending > 0 ? "text-rose-500" : "text-emerald-500"} mt-1`}>${billingPending}</h4>
                              </div>
                              <div className="w-10 h-10 bg-rose-50 dark:bg-pink-950/20 rounded-xl flex items-center justify-center text-rose-500"><CreditCard size={20} /></div>
                            </div>
                          </div>

                          {/* Payment Methods and Deposits Configured */}
                          <div className={bgCard}>
                            <div className="flex items-center space-x-2 border-b pb-4 mb-4 border-pink-500/10">
                              <CreditCard className="text-pink-500" size={20} />
                              <h3 className={`text-lg font-bold ${textTitle}`}>
                                {lang === "es" ? "Canales de Pago & Cuentas de Depósito" : "Deposit Methods & Payment Gateways"}
                              </h3>
                            </div>
                            
                            <p className="text-xs text-gray-500 leading-relaxed mb-4">
                              {lang === "es" ? "Realice sus abonos y reportes utilizando cualquiera de los canales de depósito autorizados por la Dra. Diana Rojas. Los métodos activos son gestionados directamente por administración." : "Make your deposits using any of the deposit channels authorized by Dr. Diana Rojas. Active methods are managed directly by administration."}
                            </p>

                            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                              {gateways.filter(gw => gw.enabled).map(gw => (
                                <div
                                  key={gw.id}
                                  className={`p-5 rounded-2xl border ${borderCol} ${isDark ? "bg-[#1f1618]/60 hover:bg-[#1f1618]" : "bg-rose-50/5 hover:bg-rose-50/10"} transition-all space-y-3 flex flex-col justify-between`}
                                >
                                  <div>
                                    <div className="flex items-center space-x-2 pb-2 border-b border-pink-500/5 mb-2">
                                      <div className="p-1.5 rounded-lg bg-pink-500 text-white"><CreditCard size={14} /></div>
                                      <h4 className={`text-sm font-bold ${textTitle}`}>{gw.name}</h4>
                                    </div>
                                    <p className="text-xs text-gray-600 dark:text-pink-200 leading-relaxed italic whitespace-pre-line bg-rose-50/10 dark:bg-black/10 p-3 rounded-xl border border-pink-500/5">
                                      {gw.instructions}
                                    </p>
                                  </div>

                                  {gw.link && gw.link !== "https://" && (
                                    <a
                                      href={gw.link}
                                      target="_blank"
                                      rel="noopener noreferrer"
                                      className="w-full py-2 bg-pink-500 hover:bg-pink-600 text-white rounded-xl text-xs font-bold text-center block transition-all shadow-md shadow-pink-500/5 flex items-center justify-center space-x-1"
                                    >
                                      <span>{lang === "es" ? "Ir al Enlace de Pago / Depósito" : "Go to Payment Link"}</span>
                                      <span className="text-[10px]">↗</span>
                                    </a>
                                  )}
                                </div>
                              ))}
                              {gateways.filter(gw => gw.enabled).length === 0 && (
                                <div className="col-span-1 md:col-span-2 text-center py-6 text-gray-400 italic text-xs">
                                  {lang === "es" ? "No hay pasarelas de pago habilitadas temporalmente. Por favor contacte soporte de administración." : "No payment methods configured temporarily. Please contact clinic admin."}
                                </div>
                              )}
                            </div>
                          </div>

                          {/* Invoice Records */}
                          <div className={bgCard}>
                            <div className="flex items-center space-x-2 border-b pb-4 mb-4 border-pink-500/10">
                              <DollarSign className="text-pink-500" size={20} />
                              <h3 className={`text-lg font-bold ${textTitle}`}>
                                {lang === "es" ? "Historial de Facturas de Caja y Recibos" : "Cashier Bills & Payment Invoices"}
                              </h3>
                            </div>

                            {patientInvoices.length === 0 ? (
                              <div className="py-12 text-center text-gray-400 text-xs font-bold uppercase tracking-wider">
                                No se registran facturas de caja asociadas a su nombre.
                              </div>
                            ) : (
                              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                                {patientInvoices.map(invoice => (
                                  <div
                                    key={invoice.id}
                                    className={`p-4 rounded-2xl border ${borderCol} ${isDark ? "bg-[#1f1618]" : "bg-rose-50/10"} flex justify-between items-center gap-4 hover:border-pink-500/30 transition-all`}
                                  >
                                    <div className="space-y-1">
                                      <div className="flex items-center space-x-2">
                                        <span className="px-2 py-0.5 rounded bg-pink-100 dark:bg-pink-950/30 text-pink-500 text-[9px] font-mono font-bold uppercase tracking-wider">{invoice.id.toUpperCase()}</span>
                                        <span className="text-[10px] text-gray-400 font-mono font-bold">{invoice.date}</span>
                                      </div>
                                      <h4 className={`text-sm font-bold ${textTitle}`}>{invoice.treatmentType}</h4>
                                      <p className="text-xs text-gray-500 font-semibold">Costo: ${invoice.amountTotal} • Abonado: <span className="text-emerald-500">${invoice.amountPaid}</span></p>
                                    </div>

                                    <button
                                      onClick={() => setSelectedInvoice(invoice)}
                                      className="p-3 bg-gradient-to-r from-pink-400 to-rose-500 hover:from-pink-500 hover:to-rose-600 text-white rounded-xl transition-all shadow-md shadow-pink-500/10"
                                      title="Imprimir o Ver Recibo Oficial"
                                    >
                                      <Printer size={15} />
                                    </button>
                                  </div>
                                ))}
                              </div>
                            )}
                          </div>
                        </motion.div>
                      )}

                      {/* TAB: Patient AI Companion */}
                      {patientActiveTab === "my-ai" && (
                        <motion.div
                          key="patient-ai"
                          initial={{ opacity: 0, y: 15 }}
                          animate={{ opacity: 1, y: 0 }}
                          exit={{ opacity: 0, y: -15 }}
                          className="grid grid-cols-1 lg:grid-cols-3 gap-8"
                        >
                          {/* Chat prompt selection & inputs */}
                          <div className={`${bgCard} lg:col-span-1 h-fit space-y-5`}>
                            <div className="flex items-center space-x-2 border-b pb-4 border-pink-500/10">
                              <Sparkles className="text-pink-500" size={20} />
                              <h3 className={`text-lg font-bold ${textTitle}`}>Asistente Clínico Dental</h3>
                            </div>

                            <p className="text-xs text-gray-500 leading-relaxed">
                              Realice consultas o reporte molestias clínicas. Nuestro modelo de Inteligencia Artificial le brindará orientación médica y recomendaciones preventivas inmediatas.
                            </p>

                            {/* Pre-made quick click prompts */}
                            <div className="space-y-2">
                              <span className="text-[10px] font-extrabold uppercase text-gray-400 tracking-wider">Preguntas Frecuentes Clínicas:</span>
                              {[
                                "¿Cómo debo cuidar mis carillas de disilicato de litio?",
                                "Me acaban de hacer una resina dental y tengo sensibilidad al frío, ¿es normal?",
                                "¿Cada cuánto tiempo se recomienda realizar un blanqueamiento dental láser?"
                              ].map((q, idx) => (
                                <button
                                  key={idx}
                                  type="button"
                                  onClick={() => {
                                    setPatientAiPrompt(q);
                                  }}
                                  className={`w-full text-left p-2.5 border ${borderCol} rounded-xl text-xs font-semibold text-gray-600 hover:text-pink-500 dark:text-pink-300 dark:hover:text-white hover:border-pink-500/30 hover:bg-pink-500/5 transition-all leading-normal`}
                                >
                                  {q}
                                </button>
                              ))}
                            </div>

                            <div className="space-y-3 pt-2">
                              <textarea
                                rows={3}
                                placeholder="Escriba aquí sus dudas o síntomas (Ej. Siento una molestia leve en el molar izquierdo al morder cosas calientes...)"
                                value={patientAiPrompt}
                                onChange={e => setPatientAiPrompt(e.target.value)}
                                className={`${inputStyle} resize-none`}
                              />

                              <button
                                onClick={handlePatientAiAnalyze}
                                disabled={isPatientAiLoading || !patientAiPrompt.trim()}
                                className="w-full py-3 bg-gradient-to-r from-pink-400 to-rose-500 hover:from-pink-500 hover:to-rose-600 disabled:from-gray-400 disabled:to-gray-500 text-white font-bold rounded-xl text-xs uppercase tracking-wider transition-all flex items-center justify-center space-x-2 shadow-md shadow-pink-500/10"
                              >
                                <Sparkles size={14} className={isPatientAiLoading ? "animate-pulse" : ""} />
                                <span>{isPatientAiLoading ? "Analizando con IA..." : "Consultar Clínico IA"}</span>
                              </button>
                            </div>
                          </div>

                          {/* AI Advice result panel */}
                          <div className={`${bgCard} lg:col-span-2 flex flex-col min-h-[400px]`}>
                            <div className="flex items-center space-x-2 border-b pb-4 mb-4 border-pink-500/10">
                              <Sparkles className="text-pink-500" size={20} />
                              <h3 className={`text-lg font-bold ${textTitle}`}>Ficha Clínico-Preventiva Generada</h3>
                            </div>

                            <div className="flex-grow flex flex-col justify-between">
                              {isPatientAiLoading ? (
                                <div className="flex-grow flex flex-col items-center justify-center space-y-4 py-12">
                                  <div className="w-12 h-12 border-4 border-pink-500/20 border-t-pink-500 rounded-full animate-spin" />
                                  <p className="text-xs text-gray-400 font-bold uppercase tracking-widest animate-pulse">Analizando sintomatología...</p>
                                </div>
                              ) : patientAiResponse ? (
                                <div className="space-y-4 pr-1">
                                  <div className="inline-flex items-center space-x-2 px-3 py-1 rounded-full bg-emerald-500/10 text-emerald-600 border border-emerald-500/20">
                                    <Check size={12} />
                                    <span className="text-[10px] font-bold uppercase tracking-wider font-mono">Análisis Clínico Concluido</span>
                                  </div>
                                  
                                  <div className="text-sm text-gray-700 dark:text-pink-100 font-medium leading-relaxed bg-rose-50/10 dark:bg-black/10 p-5 rounded-2xl border border-pink-500/5 whitespace-pre-line overflow-y-auto max-h-[380px] custom-scrollbar">
                                    {patientAiResponse}
                                  </div>

                                  <p className="text-[10px] text-gray-400 italic">
                                    Nota: Este análisis preventivo es orientativo. Se recomienda una consulta presencial con la Dra. Diana Rojas para una evaluación radiológica definitiva.
                                  </p>
                                </div>
                              ) : (
                                <div className="flex-grow flex flex-col items-center justify-center text-center py-12 space-y-4">
                                  <div className="w-16 h-16 bg-rose-50 dark:bg-pink-950/20 rounded-full flex items-center justify-center text-rose-300 dark:text-pink-500">
                                    <Sparkles size={32} />
                                  </div>
                                  <div className="space-y-1">
                                    <p className={`font-bold ${textTitle}`}>Asistente Clínico Listo</p>
                                    <p className="text-xs text-gray-400 max-w-sm">Escriba una pregunta clínica o seleccione una de las dudas sugeridas a la izquierda para generar su informe instantáneo con Inteligencia Artificial.</p>
                                  </div>
                                </div>
                              )}
                            </div>
                          </div>
                        </motion.div>
                      )}

                    </AnimatePresence>

                  </div>
                );
              })()
            )}

          </div>
        )}

      </main>

      {/* MODAL LIGHTBOX: INVOICE RECEIPT PRINTER PREVIEW */}
      <AnimatePresence>
        {selectedInvoice && (
          <div className="fixed inset-0 z-50 bg-black/75 backdrop-blur-md flex items-center justify-center p-4 animate-fade-in">
            <motion.div
              initial={{ scale: 0.95, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.95, opacity: 0 }}
              className="bg-white text-slate-800 rounded-3xl max-w-lg w-full shadow-2xl p-8 border border-rose-100 font-sans relative flex flex-col"
            >
              <button
                onClick={() => setSelectedInvoice(null)}
                className="absolute top-4 right-4 p-2 rounded-full hover:bg-rose-50 text-slate-400 hover:text-slate-700 transition-colors"
              >
                <XCircle size={22} />
              </button>

              {/* Receipt Header */}
              <div className="text-center space-y-1 pb-6 border-b border-rose-100">
                <div className="w-12 h-12 bg-rose-500 text-white rounded-2xl flex items-center justify-center mx-auto text-xl shadow-lg shadow-rose-200">
                  <Sparkles size={20} />
                </div>
                <h3 className="text-xl font-bold text-rose-950 font-display mt-2">DianaSRL</h3>
                <p className="text-[10px] uppercase font-mono tracking-wider font-extrabold text-pink-500">Sistema Odontológico Stetic</p>
                <p className="text-xs text-gray-400">República Dominicana • Tel: 809-555-0100</p>
              </div>

              {/* Receipt Body fields */}
              <div className="py-6 space-y-4 text-xs font-semibold">
                <div className="flex justify-between">
                  <span className="text-gray-400 uppercase tracking-wider text-[10px] font-bold">Código Factura:</span>
                  <span className="text-rose-950 font-mono font-bold">{selectedInvoice.id.toUpperCase()}</span>
                </div>
                <div className="flex justify-between">
                  <span className="text-gray-400 uppercase tracking-wider text-[10px] font-bold">Fecha Emisión:</span>
                  <span className="text-rose-950">{selectedInvoice.date}</span>
                </div>
                <div className="flex justify-between border-t border-rose-50 pt-3">
                  <span className="text-gray-400 uppercase tracking-wider text-[10px] font-bold">Cliente / Paciente:</span>
                  <span className="text-rose-950 font-bold">{selectedInvoice.patientName}</span>
                </div>
                <div className="flex justify-between">
                  <span className="text-gray-400 uppercase tracking-wider text-[10px] font-bold">Servicio / Tratamiento:</span>
                  <span className="text-rose-500 font-bold">{selectedInvoice.treatmentType}</span>
                </div>
                <div className="flex justify-between">
                  <span className="text-gray-400 uppercase tracking-wider text-[10px] font-bold">Método de Pago:</span>
                  <span className="bg-rose-100 text-rose-700 px-2 py-0.5 rounded text-[10px] font-mono font-bold uppercase">{selectedInvoice.paymentMethod}</span>
                </div>

                <div className="border-t border-rose-100 pt-4 space-y-2">
                  <div className="flex justify-between text-sm">
                    <span className="text-slate-500 font-bold">Costo Total:</span>
                    <span className="text-rose-950 font-black">${selectedInvoice.amountTotal}</span>
                  </div>
                  <div className="flex justify-between text-sm">
                    <span className="text-slate-500 font-bold">Total Abonado:</span>
                    <span className="text-emerald-600 font-black">${selectedInvoice.amountPaid}</span>
                  </div>
                  {selectedInvoice.amountPending > 0 && (
                    <div className="flex justify-between text-sm border-t border-dashed pt-2 border-slate-200">
                      <span className="text-amber-600 font-bold">Balance Pendiente:</span>
                      <span className="text-amber-600 font-black">${selectedInvoice.amountPending}</span>
                    </div>
                  )}
                </div>
              </div>

              {/* Receipt Footer */}
              <div className="border-t border-rose-100 pt-6 text-center space-y-4">
                <p className="text-[10px] text-gray-400 italic leading-relaxed">
                  "Gracias por confiar su salud dental y estética a DianaSRL. Cada sonrisa es una obra de arte diseñada con excelencia."
                </p>
                <div className="flex justify-center space-x-3">
                  <button
                    onClick={() => {
                      window.print();
                    }}
                    className="px-5 py-2.5 bg-rose-500 hover:bg-rose-600 text-white text-xs font-bold uppercase tracking-wider rounded-xl transition-all shadow-md shadow-rose-200"
                  >
                    Imprimir Factura
                  </button>
                  <button
                    onClick={() => setSelectedInvoice(null)}
                    className="px-5 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold uppercase tracking-wider rounded-xl transition-all"
                  >
                    Cerrar
                  </button>
                </div>
              </div>

            </motion.div>
          </div>
        )}
        {showPoliciesModal && (
          <div className="fixed inset-0 z-50 bg-black/75 backdrop-blur-md flex items-center justify-center p-4 animate-fade-in">
            <motion.div
              initial={{ scale: 0.95, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.95, opacity: 0 }}
              className={`rounded-3xl max-w-2xl w-full shadow-2xl p-6 border ${borderCol} ${isDark ? "bg-[#160f11] text-pink-100" : "bg-white text-slate-800"} font-sans relative flex flex-col max-h-[90vh]`}
            >
              <button
                onClick={() => setShowPoliciesModal(false)}
                className="absolute top-4 right-4 p-2 rounded-full hover:bg-rose-500/10 text-gray-400 hover:text-rose-500 transition-colors flex items-center justify-center"
              >
                <XCircle size={22} />
              </button>

              <div className="flex items-center space-x-2 pb-4 mb-4 border-b border-pink-500/10">
                <ShieldCheck className="text-pink-500" size={24} />
                <div>
                  <h3 className={`text-lg font-bold ${textTitle}`}>
                    {lang === "es" ? "Centro Legal de Políticas" : "Legal & Policy Center"}
                  </h3>
                  <p className="text-[10px] text-gray-400 uppercase font-mono font-bold tracking-wider">
                    {clinicName} • {lang === "es" ? "Transparencia Dental" : "Dental Transparency"}
                  </p>
                </div>
              </div>

              {/* Tab Selector inside Modal */}
              <div className="flex border-b border-pink-500/10 mb-4 text-xs font-bold uppercase tracking-wider">
                <button
                  type="button"
                  onClick={() => setActivePoliciesTab("terms")}
                  className={`pb-2.5 px-4 border-b-2 transition-all ${
                    activePoliciesTab === "terms"
                      ? "border-pink-500 text-pink-500 font-extrabold"
                      : "border-transparent text-gray-400 hover:text-gray-500"
                  }`}
                >
                  {lang === "es" ? "Términos y Condiciones" : "Terms & Conditions"}
                </button>
                <button
                  type="button"
                  onClick={() => setActivePoliciesTab("privacy")}
                  className={`pb-2.5 px-4 border-b-2 transition-all ${
                    activePoliciesTab === "privacy"
                      ? "border-pink-500 text-pink-500 font-extrabold"
                      : "border-transparent text-gray-400 hover:text-gray-500"
                  }`}
                >
                  {lang === "es" ? "Políticas de Privacidad" : "Privacy Policy"}
                </button>
              </div>

              {/* Policy Body Content */}
              <div className="flex-grow overflow-y-auto pr-2 custom-scrollbar text-xs leading-relaxed text-gray-500 dark:text-pink-100/70 font-sans space-y-4 max-h-[50vh] whitespace-pre-wrap p-3 bg-pink-500/[0.02] rounded-2xl border border-pink-500/5">
                {activePoliciesTab === "terms" ? termsAndConditions : privacyPolicy}
              </div>

              <div className="pt-4 border-t border-pink-500/10 mt-4 flex justify-end">
                <button
                  type="button"
                  onClick={() => setShowPoliciesModal(false)}
                  className={`px-5 py-2.5 bg-gradient-to-r ${activeThemeObj.gradient} text-white text-xs font-bold uppercase tracking-wider rounded-xl transition-all shadow-md ${activeThemeObj.shadow}`}
                >
                  {lang === "es" ? "Entendido y Cerrar" : "Understood & Close"}
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Brand Footer */}
      <footer className={`max-w-7xl mx-auto px-6 py-12 mt-16 border-t ${borderCol} flex flex-col md:flex-row items-center justify-between gap-6`}>
        <div className="space-y-1">
          <div className="flex items-center space-x-2">
            <span className={`w-2.5 h-2.5 rounded-full ${activeThemeObj.accentText} animate-pulse bg-current`} />
            <span className={`text-xs font-bold uppercase tracking-wider ${activeThemeObj.accentText}`}>{clinicName} — {clinicSlogan}</span>
          </div>
          <p className="text-[10px] text-gray-400 uppercase tracking-widest font-mono">
            &copy; 2026 {clinicName.toUpperCase()} CLINIC SYSTEMS. ALL RIGHTS RESERVED.
          </p>
        </div>

        <div className="flex flex-wrap gap-6 text-xs font-bold uppercase tracking-widest text-gray-500">
          <button 
            onClick={() => {
              setShowPoliciesModal(true);
              setActivePoliciesTab("terms");
            }} 
            className="hover:text-pink-500 transition-colors"
          >
            {lang === "es" ? "Términos y Condiciones" : "Terms & Conditions"}
          </button>
          <button 
            onClick={() => {
              setShowPoliciesModal(true);
              setActivePoliciesTab("privacy");
            }} 
            className="hover:text-pink-500 transition-colors"
          >
            {lang === "es" ? "Políticas de Privacidad" : "Privacy Policy"}
          </button>
        </div>
      </footer>
    </div>
  );
}
