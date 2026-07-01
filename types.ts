export interface Patient {
  id: string;
  name: string;
  phone: string;
  email: string;
  dni: string;
  dob: string;
  notes: string;
  registeredAt: string;
  allergies?: string;
  medicalConditions?: string;
  consultationReason?: string;
  dentalInsurance?: string;
  preferredContact?: string;
  address?: string;
}

export interface Appointment {
  id: string;
  patientId: string;
  patientName: string;
  dentistName: string;
  date: string;
  time: string;
  treatmentType: string;
  notes: string;
  status: "Scheduled" | "Completed" | "Cancelled";
}

export interface ToothCondition {
  toothNumber: number;
  status: "Healthy" | "Caries" | "Endo" | "Missing" | "Crown" | "Filling";
  notes?: string;
  updatedAt: string;
}

export interface PatientChart {
  patientId: string;
  conditions: ToothCondition[];
}

export interface PaymentRecord {
  id: string;
  patientId: string;
  patientName: string;
  treatmentType: string;
  amountTotal: number;
  amountPaid: number;
  amountPending: number;
  paymentMethod: "Cash" | "Card" | "Transfer";
  paymentStatus: "Pending" | "Partial" | "Paid";
  date: string;
}

export interface SyncLog {
  id: string;
  timestamp: string;
  action: string;
  status: "Success" | "Error" | "Pending";
  message: string;
}

export interface AdminUser {
  id: string;
  username: string;
  name: string;
  passwordHash: string; // Stored simply in this client/server context
  role: "master" | "admin";
  createdAt: string;
}

export interface DoctorAvailability {
  id: string;
  doctorName: string; // e.g., "Dra. Diana Rojas" or other registered admins
  date: string; // "YYYY-MM-DD"
  slots: string[]; // e.g., ["08:00", "09:00", "10:00", "11:00", "14:00", "15:00", "16:00"]
}

export interface PaymentMethodConfig {
  id: string;
  name: string;
  enabled: boolean;
  instructions: string;
  link: string;
}



