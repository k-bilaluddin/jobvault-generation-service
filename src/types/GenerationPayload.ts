export interface SkillRow {
  label: string;
  value: string;
}

export interface RolePayload {
  /**
   * Stable identifier for the role — must be one of:
   *   "calvergy" | "senior_baris" | "developer_baris" | "junior_baris"
   * Company, title, dates are fixed in the template; only bullets change per JD.
   */
  id: string;
  bullets: string[];
}

export interface GenerationPayload {
  company: string;
  role: string;
  jdSource: string;
  headline: string;
  summary: string;
  skills: SkillRow[];
  roles: RolePayload[];
  recipient: string;
  coverLetterParagraphs: string[];
  compatibilityScore: number;
  strengths: string[];
  gaps: string[];
  tailoringNotes: string;
}
