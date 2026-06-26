import React, { useState, useMemo, useRef, useEffect } from "react";

// Google Fonts — Inter
if (typeof document !== 'undefined' && !document.getElementById('inter-font')) {
  const link = document.createElement('link');
  link.id = 'inter-font';
  link.rel = 'stylesheet';
  link.href = 'https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap';
  document.head.appendChild(link);
}
import * as XLSX from "xlsx";

// ─── Design Tokens ────────────────────────────────────────────────────────────
const C = {
  navBg:       "#0C1220",
  navText:     "#8B9BB4",
  navActive:   "#F8FAFC",
  navHover:    "#161F2E",
  accent:      "#1E3A5F",
  accentLight: "#2563EB",
  contentBg:   "#F4F6F9",
  white:       "#FFFFFF",
  border:      "#E8EDF2",
  borderDark:  "#CDD5E0",
  textPrimary: "#0D1B2E",
  textMuted:   "#5A6A7E",
  textLight:   "#96A3B5",
  success:     "#15803D",
  successBg:   "#DCFCE7",
  neutral:     "#64748B",
  neutralBg:   "#F1F5F9",
  danger:      "#DC2626",
  dangerBg:    "#FEE2E2",
  warning:     "#B45309",
  warningBg:   "#FEF3C7",
  tableHover:  "#EEF4FF",
  tableStripe: "#FAFBFD",
};

// ─── CSI MasterFormat Data ────────────────────────────────────────────────────
const CSI_STRUCTURE = [
  // ── FACILITY CONSTRUCTION SUBGROUP ──────────────────────────────────────────
  {
    code: "02", label: "Existing Conditions",
    subdivisions: [
      { code: "02 20 00", label: "Assessment" },
      { code: "02 30 00", label: "Subsurface Investigation" },
      { code: "02 40 00", label: "Demolition & Structure Moving" },
      { code: "02 50 00", label: "Site Remediation" },
      { code: "02 60 00", label: "Contaminated Site Material Removal" },
      { code: "02 70 00", label: "Water Remediation" },
      { code: "02 80 00", label: "Facility Remediation" },
    ]
  },
  {
    code: "03", label: "Concrete",
    subdivisions: [
      { code: "03 10 00", label: "Concrete Forming & Accessories" },
      { code: "03 20 00", label: "Concrete Reinforcing" },
      { code: "03 30 00", label: "Cast-in-Place Concrete" },
      { code: "03 40 00", label: "Precast Concrete" },
      { code: "03 50 00", label: "Cast Decks & Underlayment" },
      { code: "03 60 00", label: "Grouting" },
      { code: "03 70 00", label: "Mass Concrete" },
      { code: "03 80 00", label: "Concrete Cutting & Boring" },
    ]
  },
  {
    code: "04", label: "Masonry",
    subdivisions: [
      { code: "04 20 00", label: "Unit Masonry" },
      { code: "04 40 00", label: "Stone Assemblies" },
      { code: "04 50 00", label: "Refractory Masonry" },
      { code: "04 60 00", label: "Corrosion-Resistant Masonry" },
      { code: "04 70 00", label: "Manufactured Masonry" },
    ]
  },
  {
    code: "05", label: "Metals",
    subdivisions: [
      { code: "05 10 00", label: "Structural Metal Framing" },
      { code: "05 20 00", label: "Metal Joists" },
      { code: "05 30 00", label: "Metal Decking" },
      { code: "05 40 00", label: "Cold-Formed Metal Framing" },
      { code: "05 50 00", label: "Metal Fabrications" },
      { code: "05 70 00", label: "Decorative Metal" },
    ]
  },
  {
    code: "06", label: "Wood, Plastics & Composites",
    subdivisions: [
      { code: "06 10 00", label: "Rough Carpentry" },
      { code: "06 20 00", label: "Finish Carpentry" },
      { code: "06 40 00", label: "Architectural Woodwork" },
      { code: "06 50 00", label: "Structural Plastics" },
      { code: "06 60 00", label: "Plastic Fabrications" },
      { code: "06 70 00", label: "Structural Composites" },
      { code: "06 80 00", label: "Composite Fabrications" },
    ]
  },
  {
    code: "07", label: "Thermal & Moisture Protection",
    subdivisions: [
      { code: "07 10 00", label: "Dampproofing & Waterproofing" },
      { code: "07 20 00", label: "Thermal Protection" },
      { code: "07 25 00", label: "Weather Barriers" },
      { code: "07 30 00", label: "Steep Slope Roofing" },
      { code: "07 40 00", label: "Roofing & Siding Panels" },
      { code: "07 50 00", label: "Membrane Roofing" },
      { code: "07 60 00", label: "Flashing & Sheet Metal" },
      { code: "07 70 00", label: "Roof & Wall Specialties & Accessories" },
      { code: "07 80 00", label: "Fire & Smoke Protection" },
      { code: "07 90 00", label: "Joint Protection" },
    ]
  },
  {
    code: "08", label: "Openings",
    subdivisions: [
      { code: "08 10 00", label: "Doors & Frames" },
      { code: "08 30 00", label: "Specialty Doors & Frames" },
      { code: "08 40 00", label: "Entrances, Storefronts & Curtain Walls" },
      { code: "08 50 00", label: "Windows" },
      { code: "08 60 00", label: "Roof Windows & Skylights" },
      { code: "08 70 00", label: "Hardware" },
      { code: "08 80 00", label: "Glazing" },
      { code: "08 90 00", label: "Louvers & Vents" },
    ]
  },
  {
    code: "09", label: "Finishes",
    subdivisions: [
      { code: "09 20 00", label: "Plaster & Gypsum Board" },
      { code: "09 30 00", label: "Tiling" },
      { code: "09 50 00", label: "Ceilings" },
      { code: "09 60 00", label: "Flooring" },
      { code: "09 70 00", label: "Wall Finishes" },
      { code: "09 80 00", label: "Acoustic Treatment" },
      { code: "09 90 00", label: "Painting & Coating" },
    ]
  },
  {
    code: "10", label: "Specialties",
    subdivisions: [
      { code: "10 10 00", label: "Information Specialties" },
      { code: "10 20 00", label: "Interior Specialties" },
      { code: "10 30 00", label: "Fireplaces & Stoves" },
      { code: "10 40 00", label: "Safety Specialties" },
      { code: "10 50 00", label: "Storage Specialties" },
      { code: "10 70 00", label: "Exterior Specialties" },
      { code: "10 80 00", label: "Other Specialties" },
    ]
  },
  {
    code: "11", label: "Equipment",
    subdivisions: [
      { code: "11 10 00", label: "Vehicle & Pedestrian Equipment" },
      { code: "11 15 00", label: "Security, Detention & Banking Equipment" },
      { code: "11 20 00", label: "Commercial Equipment" },
      { code: "11 30 00", label: "Residential Equipment" },
      { code: "11 40 00", label: "Foodservice Equipment" },
      { code: "11 50 00", label: "Educational & Scientific Equipment" },
      { code: "11 60 00", label: "Entertainment Equipment" },
      { code: "11 65 00", label: "Athletic & Recreational Equipment" },
      { code: "11 70 00", label: "Healthcare Equipment" },
      { code: "11 80 00", label: "Collection & Disposal Equipment" },
      { code: "11 90 00", label: "Other Equipment" },
    ]
  },
  {
    code: "12", label: "Furnishings",
    subdivisions: [
      { code: "12 10 00", label: "Art" },
      { code: "12 20 00", label: "Window Treatments" },
      { code: "12 30 00", label: "Casework" },
      { code: "12 40 00", label: "Furnishings & Accessories" },
      { code: "12 50 00", label: "Furniture" },
      { code: "12 60 00", label: "Multiple Seating" },
      { code: "12 90 00", label: "Other Furnishings" },
    ]
  },
  {
    code: "13", label: "Special Construction",
    subdivisions: [
      { code: "13 10 00", label: "Special Facility Components" },
      { code: "13 20 00", label: "Special Purpose Rooms" },
      { code: "13 30 00", label: "Special Structures" },
      { code: "13 40 00", label: "Integrated Construction" },
      { code: "13 50 00", label: "Special Instrumentation" },
    ]
  },
  {
    code: "14", label: "Conveying Systems",
    subdivisions: [
      { code: "14 10 00", label: "Dumbwaiters" },
      { code: "14 20 00", label: "Elevators" },
      { code: "14 30 00", label: "Escalators & Moving Walks" },
      { code: "14 40 00", label: "Lifts" },
      { code: "14 70 00", label: "Turntables" },
      { code: "14 80 00", label: "Scaffolding" },
      { code: "14 90 00", label: "Other Conveying Equipment" },
    ]
  },
  // ── FACILITY SERVICES SUBGROUP ───────────────────────────────────────────────
  {
    code: "21", label: "Fire Suppression",
    subdivisions: [
      { code: "21 10 00", label: "Water-Based Fire-Suppression Systems" },
      { code: "21 20 00", label: "Fire-Extinguishing Systems" },
      { code: "21 30 00", label: "Fire Pumps" },
      { code: "21 40 00", label: "Fire-Suppression Water Storage" },
    ]
  },
  {
    code: "22", label: "Plumbing",
    subdivisions: [
      { code: "22 10 00", label: "Plumbing Piping & Pumps" },
      { code: "22 30 00", label: "Plumbing Equipment" },
      { code: "22 40 00", label: "Plumbing Fixtures" },
      { code: "22 50 00", label: "Pool & Fountain Plumbing Systems" },
      { code: "22 60 00", label: "Gas & Vacuum Systems for Laboratory & Healthcare Facilities" },
    ]
  },
  {
    code: "23", label: "Heating, Ventilating & Air-Conditioning (HVAC)",
    subdivisions: [
      { code: "23 10 00", label: "Facility Fuel Systems" },
      { code: "23 20 00", label: "HVAC Piping & Pumps" },
      { code: "23 30 00", label: "HVAC Air Distribution" },
      { code: "23 40 00", label: "HVAC Air Cleaning Devices" },
      { code: "23 50 00", label: "Central Heating Equipment" },
      { code: "23 60 00", label: "Central Cooling Equipment" },
      { code: "23 70 00", label: "Central HVAC Equipment" },
      { code: "23 80 00", label: "Decentralized HVAC Equipment" },
    ]
  },
  {
    code: "25", label: "Integrated Automation",
    subdivisions: [
      { code: "25 10 00", label: "Integrated Automation Network Equipment" },
      { code: "25 30 00", label: "Integrated Automation Instrumentation & Terminal Devices" },
      { code: "25 50 00", label: "Integrated Automation Facility Controls" },
      { code: "25 90 00", label: "Integrated Automation Control Sequences" },
    ]
  },
  {
    code: "26", label: "Electrical",
    subdivisions: [
      { code: "26 10 00", label: "Medium-Voltage Electrical Distribution" },
      { code: "26 20 00", label: "Low-Voltage Electrical Transmission" },
      { code: "26 30 00", label: "Facility Electrical Power Generating & Storage Equipment" },
      { code: "26 40 00", label: "Electrical & Cathodic Protection" },
      { code: "26 50 00", label: "Lighting" },
    ]
  },
  {
    code: "27", label: "Communications",
    subdivisions: [
      { code: "27 10 00", label: "Structured Cabling" },
      { code: "27 20 00", label: "Data Communications" },
      { code: "27 30 00", label: "Voice Communications" },
      { code: "27 40 00", label: "Audio-Video Communications" },
      { code: "27 50 00", label: "Distributed Communications & Monitoring Systems" },
      { code: "27 60 00", label: "Wireless Transceivers" },
    ]
  },
  {
    code: "28", label: "Electronic Safety & Security",
    subdivisions: [
      { code: "28 10 00", label: "Electronic Access Control & Intrusion Detection" },
      { code: "28 20 00", label: "Electronic Surveillance" },
      { code: "28 30 00", label: "Electronic Detection & Alarm" },
      { code: "28 40 00", label: "Electronic Monitoring & Control" },
    ]
  },
  // ── SITE AND INFRASTRUCTURE SUBGROUP ────────────────────────────────────────
  {
    code: "31", label: "Earthwork",
    subdivisions: [
      { code: "31 10 00", label: "Site Clearing" },
      { code: "31 20 00", label: "Earth Moving" },
      { code: "31 30 00", label: "Earthwork Methods" },
      { code: "31 40 00", label: "Shoring & Underpinning" },
      { code: "31 50 00", label: "Excavation Support & Protection" },
      { code: "31 60 00", label: "Special Foundations & Load-Bearing Elements" },
      { code: "31 70 00", label: "Tunneling & Mining" },
    ]
  },
  {
    code: "32", label: "Exterior Improvements",
    subdivisions: [
      { code: "32 10 00", label: "Bases, Ballasts & Paving" },
      { code: "32 30 00", label: "Site Improvements" },
      { code: "32 70 00", label: "Wetlands" },
      { code: "32 80 00", label: "Irrigation" },
      { code: "32 90 00", label: "Planting" },
    ]
  },
  {
    code: "33", label: "Utilities",
    subdivisions: [
      { code: "33 10 00", label: "Water Utilities" },
      { code: "33 20 00", label: "Wells" },
      { code: "33 30 00", label: "Sanitary Sewerage Utilities" },
      { code: "33 40 00", label: "Storm Drainage Utilities" },
      { code: "33 50 00", label: "Fuel Distribution Utilities" },
      { code: "33 60 00", label: "Hydronic & Steam Energy Utilities" },
      { code: "33 70 00", label: "Electrical Utilities" },
      { code: "33 80 00", label: "Communications Utilities" },
    ]
  },
];

// ─── Mock Cost Items ──────────────────────────────────────────────────────────
let NEXT_ID = 1000;
const generateId = () => `CI-${++NEXT_ID}`;

const MOCK_ITEMS = [
  // 03 30 00 Cast-in-Place Concrete
  { id: "CI-001", org_id: "org-1", csi_code: "03 30 00", description: "Slab on Grade – 4\" w/ WWF", unit: "SF", mat_cost: 210, lab_cost: 185, equip_cost: 40, region: "Dallas–Fort Worth", source: "Actuals", updated_by: "J. Rourke", updated_at: "2025-06-10", notes: "Includes vapor barrier, wire mesh, and broom finish." },
  { id: "CI-002", org_id: "org-1", csi_code: "03 30 00", description: "Slab on Grade – 6\" w/ Rebar", unit: "SF", mat_cost: 310, lab_cost: 240, equip_cost: 55, region: "Dallas–Fort Worth", source: "Bid", updated_by: "M. Chen", updated_at: "2025-05-22", notes: "" },
  { id: "CI-003", org_id: "org-1", csi_code: "03 30 00", description: "Elevated Slab – PT", unit: "SF", mat_cost: 520, lab_cost: 380, equip_cost: 95, region: "Houston", source: "Estimate", updated_by: "J. Rourke", updated_at: "2025-04-15", notes: "Post-tension system; verify PT supplier lead times." },
  { id: "CI-004", org_id: "org-1", csi_code: "03 30 00", description: "Spread Footing – 3000 PSI", unit: "CY", mat_cost: 18500, lab_cost: 9200, equip_cost: 2100, region: "Dallas–Fort Worth", source: "Actuals", updated_by: "J. Rourke", updated_at: "2025-06-01", notes: "" },
  { id: "CI-005", org_id: "org-1", csi_code: "03 30 00", description: "Grade Beam – 4000 PSI", unit: "LF", mat_cost: 8400, lab_cost: 5600, equip_cost: 1200, region: "Dallas–Fort Worth", source: "Actuals", updated_by: "S. Patel", updated_at: "2025-03-18", notes: "" },

  // 03 20 00 Concrete Reinforcing
  { id: "CI-006", org_id: "org-1", csi_code: "03 20 00", description: "#4 Rebar – Epoxy Coated", unit: "TON", mat_cost: 142000, lab_cost: 48000, equip_cost: 0, region: "Dallas–Fort Worth", source: "Bid", updated_by: "M. Chen", updated_at: "2025-06-18", notes: "Pricing valid through Q3 2025. Verify with supplier." },
  { id: "CI-007", org_id: "org-1", csi_code: "03 20 00", description: "#5 Rebar – Black", unit: "TON", mat_cost: 118000, lab_cost: 44000, equip_cost: 0, region: "Dallas–Fort Worth", source: "Bid", updated_by: "M. Chen", updated_at: "2025-06-18", notes: "" },
  { id: "CI-008", org_id: "org-1", csi_code: "03 20 00", description: "Welded Wire Reinforcement", unit: "SF", mat_cost: 38, lab_cost: 22, equip_cost: 0, region: "Dallas–Fort Worth", source: "RS Means", updated_by: "J. Rourke", updated_at: "2025-01-10", notes: "RS Means 2025 DFW adjustment applied." },

  // 03 10 00 Concrete Forming
  { id: "CI-009", org_id: "org-1", csi_code: "03 10 00", description: "Wall Form – Single Use Ply", unit: "SFCA", mat_cost: 145, lab_cost: 210, equip_cost: 30, region: "Dallas–Fort Worth", source: "Actuals", updated_by: "S. Patel", updated_at: "2025-02-20", notes: "" },
  { id: "CI-010", org_id: "org-1", csi_code: "03 10 00", description: "Column Form – Round Tube", unit: "LF", mat_cost: 420, lab_cost: 380, equip_cost: 0, region: "Dallas–Fort Worth", source: "Actuals", updated_by: "S. Patel", updated_at: "2025-02-20", notes: "" },

  // 05 10 00 Structural Metal Framing
  { id: "CI-011", org_id: "org-1", csi_code: "05 10 00", description: "Structural Steel – A992 W-Shape", unit: "TON", mat_cost: 248000, lab_cost: 112000, equip_cost: 38000, region: "Dallas–Fort Worth", source: "Bid", updated_by: "J. Rourke", updated_at: "2025-06-05", notes: "Includes shop drawings, fab, and erection." },
  { id: "CI-012", org_id: "org-1", csi_code: "05 10 00", description: "HSS Columns – 6x6", unit: "TON", mat_cost: 265000, lab_cost: 98000, equip_cost: 32000, region: "Dallas–Fort Worth", source: "Bid", updated_by: "J. Rourke", updated_at: "2025-06-05", notes: "" },
  { id: "CI-013", org_id: "org-1", csi_code: "05 10 00", description: "Anchor Bolts – Cast-in-Place", unit: "EA", mat_cost: 4800, lab_cost: 1200, equip_cost: 0, region: "Houston", source: "Actuals", updated_by: "M. Chen", updated_at: "2025-05-01", notes: "" },

  // 05 30 00 Metal Decking
  { id: "CI-014", org_id: "org-1", csi_code: "05 30 00", description: '3" Composite Deck – 20GA', unit: "SF", mat_cost: 385, lab_cost: 145, equip_cost: 25, region: "Dallas–Fort Worth", source: "Bid", updated_by: "J. Rourke", updated_at: "2025-05-14", notes: "" },
  { id: "CI-015", org_id: "org-1", csi_code: "05 30 00", description: '1.5" Roof Deck – 22GA', unit: "SF", mat_cost: 280, lab_cost: 120, equip_cost: 20, region: "Dallas–Fort Worth", source: "Bid", updated_by: "J. Rourke", updated_at: "2025-05-14", notes: "" },

  // 09 20 00 Plaster and Gypsum Board
  { id: "CI-016", org_id: "org-1", csi_code: "09 20 00", description: "5/8\" Type X GWB – Single Layer", unit: "SF", mat_cost: 62, lab_cost: 88, equip_cost: 0, region: "Dallas–Fort Worth", source: "Actuals", updated_by: "S. Patel", updated_at: "2025-04-08", notes: "" },
  { id: "CI-017", org_id: "org-1", csi_code: "09 20 00", description: "Metal Stud Framing – 3-5/8\" 20GA", unit: "SF", mat_cost: 48, lab_cost: 110, equip_cost: 0, region: "Dallas–Fort Worth", source: "Actuals", updated_by: "S. Patel", updated_at: "2025-04-08", notes: "Wall assembly only; GWB priced separately." },

  // 09 60 00 Flooring
  { id: "CI-018", org_id: "org-1", csi_code: "09 60 00", description: "LVT Flooring – 12mil Wear Layer", unit: "SF", mat_cost: 320, lab_cost: 145, equip_cost: 0, region: "Dallas–Fort Worth", source: "Bid", updated_by: "M. Chen", updated_at: "2025-06-12", notes: "" },
  { id: "CI-019", org_id: "org-1", csi_code: "09 60 00", description: "Carpet Tile – 24x24 Modular", unit: "SY", mat_cost: 2800, lab_cost: 650, equip_cost: 0, region: "Dallas–Fort Worth", source: "Bid", updated_by: "M. Chen", updated_at: "2025-06-12", notes: "" },

  // 26 20 00 Electrical
  { id: "CI-020", org_id: "org-1", csi_code: "26 20 00", description: "Conduit – 3/4\" EMT in Slab", unit: "LF", mat_cost: 420, lab_cost: 380, equip_cost: 0, region: "Dallas–Fort Worth", source: "Actuals", updated_by: "J. Rourke", updated_at: "2025-03-30", notes: "" },
  { id: "CI-021", org_id: "org-1", csi_code: "26 20 00", description: "200A Panel – 42 Circuit", unit: "EA", mat_cost: 186000, lab_cost: 48000, equip_cost: 0, region: "Dallas–Fort Worth", source: "Bid", updated_by: "J. Rourke", updated_at: "2025-06-08", notes: "" },

  // 31 20 00 Earthwork
  { id: "CI-022", org_id: "org-1", csi_code: "31 20 00", description: "Bulk Excavation – Import Haul", unit: "CY", mat_cost: 0, lab_cost: 1850, equip_cost: 2400, region: "Dallas–Fort Worth", source: "Actuals", updated_by: "J. Rourke", updated_at: "2025-05-28", notes: "Rate includes operator, fuel, and disposal." },
  { id: "CI-023", org_id: "org-1", csi_code: "31 20 00", description: "Structural Fill – Compact & Test", unit: "CY", mat_cost: 2800, lab_cost: 1200, equip_cost: 800, region: "Dallas–Fort Worth", source: "Actuals", updated_by: "S. Patel", updated_at: "2025-04-22", notes: "" },

  // 22 40 00 Plumbing Fixtures
  { id: "CI-024", org_id: "org-1", csi_code: "22 40 00", description: "Floor Drain – Cast Iron Body", unit: "EA", mat_cost: 38000, lab_cost: 22000, equip_cost: 0, region: "Dallas–Fort Worth", source: "Estimate", updated_by: "M. Chen", updated_at: "2025-02-14", notes: "" },

  // 07 50 00 Membrane Roofing
  { id: "CI-025", org_id: "org-1", csi_code: "07 50 00", description: "TPO Membrane – 60mil Fully Adhered", unit: "SF", mat_cost: 245, lab_cost: 180, equip_cost: 15, region: "Dallas–Fort Worth", source: "Bid", updated_by: "J. Rourke", updated_at: "2025-06-16", notes: "Includes insulation board; R-25 minimum." },
];

// ─── Mock History ─────────────────────────────────────────────────────────────
const MOCK_HISTORY = {
  "CI-001": [
    { changed_at: "2025-06-10T14:22:00Z", changed_by: "J. Rourke", snapshot: { mat_cost: 210, lab_cost: 185, equip_cost: 40, source: "Actuals", notes: "Updated to reflect June actuals." } },
    { changed_at: "2025-03-02T09:10:00Z", changed_by: "S. Patel",  snapshot: { mat_cost: 195, lab_cost: 178, equip_cost: 40, source: "Estimate", notes: "" } },
    { changed_at: "2024-11-15T16:45:00Z", changed_by: "J. Rourke", snapshot: { mat_cost: 182, lab_cost: 170, equip_cost: 38, source: "RS Means", notes: "Seeded from RS Means 2024." } },
  ],
  "CI-006": [
    { changed_at: "2025-06-18T11:05:00Z", changed_by: "M. Chen",   snapshot: { mat_cost: 142000, lab_cost: 48000, equip_cost: 0, source: "Bid", notes: "Pricing valid through Q3 2025." } },
    { changed_at: "2025-01-20T08:30:00Z", changed_by: "J. Rourke", snapshot: { mat_cost: 128000, lab_cost: 45000, equip_cost: 0, source: "Estimate", notes: "" } },
  ],
  "CI-011": [
    { changed_at: "2025-06-05T15:00:00Z", changed_by: "J. Rourke", snapshot: { mat_cost: 248000, lab_cost: 112000, equip_cost: 38000, source: "Bid", notes: "Includes shop drawings, fab, and erection." } },
    { changed_at: "2025-02-10T10:20:00Z", changed_by: "M. Chen",   snapshot: { mat_cost: 230000, lab_cost: 105000, equip_cost: 35000, source: "Estimate", notes: "" } },
    { changed_at: "2024-09-01T13:00:00Z", changed_by: "J. Rourke", snapshot: { mat_cost: 218000, lab_cost: 98000,  equip_cost: 32000, source: "RS Means", notes: "RS Means baseline." } },
  ],
};

// ─── Constants ────────────────────────────────────────────────────────────────
const UNITS = ["SF", "SY", "LF", "CY", "EA", "TON", "SFCA", "LS", "MBF", "GAL", "HR"];
const SOURCES = ["Bid", "Estimate", "Actuals", "RS Means", "Other"];
const NAV_ITEMS = [
  { label: "Dashboard",       key: "dashboard" },
  { label: "Cost Database",   key: "cost-database" },
  { label: "Estimates",       key: "estimates" },
  { label: "Bid Comparisons", key: "bid-comparisons" },
];
const MOCK_ORG  = { name: "Meridian Construction Group", plan: "Enterprise" };
const MOCK_USER = { name: "James Rourke", role: "Preconstruction Manager", initials: "JR" };

// ─── Bulk Import Constants ────────────────────────────────────────────────────
const IMPORT_TEMPLATE_COLS = [
  "csi_code", "description", "unit", "mat_cost", "lab_cost",
  "equip_cost", "region", "source", "notes"
];

const IMPORT_TEMPLATE_ROWS = [
  ["03 30 00", "Slab on Grade – 4-inch w/ WWF", "SF", "2.10", "1.85", "0.40", "Dallas–Fort Worth", "Actuals", "Includes vapor barrier and broom finish"],
  ["05 10 00", "Structural Steel – A992 W-Shape", "TON", "2480.00", "1120.00", "380.00", "Dallas–Fort Worth", "Bid", "Includes fab and erection"],
  ["09 60 00", "LVT Flooring – 12mil Wear Layer", "SF", "3.20", "1.45", "0.00", "Houston", "Bid", ""],
];

const VALID_CSI_CODES = new Set(
  CSI_STRUCTURE.flatMap(div => div.subdivisions.map(s => s.code))
);
const VALID_UNITS   = new Set(["SF","SY","LF","CY","EA","TON","SFCA","LS","MBF","GAL","HR"]);
const VALID_SOURCES = new Set(["Bid","Estimate","Actuals","RS Means","Other"]);

function downloadTemplate() {
  const header = IMPORT_TEMPLATE_COLS.join(",");
  const rows   = IMPORT_TEMPLATE_ROWS.map(r =>
    r.map(cell => cell.includes(",") ? `"${cell}"` : cell).join(",")
  );
  const csv    = [header, ...rows].join("\n");
  const blob   = new Blob([csv], { type: "text/csv" });
  const url    = URL.createObjectURL(blob);
  const a      = document.createElement("a");
  a.href       = url;
  a.download   = "meridian_cost_import_template.csv";
  a.click();
  URL.revokeObjectURL(url);
}

function parseCSV(text) {
  const lines = text.trim().split("\n").map(l => l.trim()).filter(Boolean);
  if (lines.length < 2) return { error: "File appears empty or has no data rows." };
  const headers = lines[0].split(",").map(h => h.replace(/^"|"$/g, "").trim());
  const missing = IMPORT_TEMPLATE_COLS.filter(c => !headers.includes(c));
  if (missing.length) return { error: `Missing required columns: ${missing.join(", ")}` };
  const rows = [];
  for (let i = 1; i < lines.length; i++) {
    // Basic CSV parse (handles quoted fields)
    const cells = [];
    let cur = "", inQ = false;
    for (const ch of lines[i]) {
      if (ch === '"') { inQ = !inQ; }
      else if (ch === "," && !inQ) { cells.push(cur.trim()); cur = ""; }
      else cur += ch;
    }
    cells.push(cur.trim());
    const row = {};
    headers.forEach((h, idx) => { row[h] = cells[idx] || ""; });
    rows.push(row);
  }
  return { headers, rows };
}

function validateRow(row, existingItems) {
  const errors = [];
  const warnings = [];
  if (!row.csi_code?.trim())                errors.push("csi_code required");
  else if (!VALID_CSI_CODES.has(row.csi_code.trim())) errors.push(`Unknown CSI code: ${row.csi_code}`);
  if (!row.description?.trim())             errors.push("description required");
  if (!row.unit?.trim())                    errors.push("unit required");
  else if (!VALID_UNITS.has(row.unit.trim().toUpperCase())) warnings.push(`Unrecognized unit: ${row.unit}`);
  if (row.source && !VALID_SOURCES.has(row.source.trim())) warnings.push(`Unrecognized source: ${row.source}`);
  const mat  = parseFloat(row.mat_cost)  || 0;
  const lab  = parseFloat(row.lab_cost)  || 0;
  const equip= parseFloat(row.equip_cost)|| 0;
  if (mat + lab + equip === 0)              warnings.push("All costs are zero");
  // Conflict detection: same csi_code + description (case-insensitive)
  const conflict = existingItems.find(i =>
    i.csi_code === row.csi_code?.trim() &&
    i.description.toLowerCase() === row.description?.trim().toLowerCase()
  );
  return { errors, warnings, conflict: conflict || null };
}


const QUOTE_STATUSES = ["Received", "Reviewed", "Approved", "Rejected"];

let NEXT_QUOTE_ID = 100;
const generateQuoteId = () => `Q-${++NEXT_QUOTE_ID}`;

const MOCK_QUOTES = [
  {
    id: "Q-001",
    org_id: "org-1",
    company: "Lone Star Concrete LLC",
    trade: "Concrete",
    contact_name: "Mike Torres",
    email: "mtorres@lonestar.com",
    phone: "214-555-0182",
    market: "Dallas–Fort Worth",
    cost_item_id: "CI-001",
    csi_code: "03 30 00",
    csi_label: "Cast-in-Place Concrete",
    description: "SOG 4-inch w/ vapor barrier, WWF, broom finish",
    unit: "SF",
    unit_cost: 875,   // cents
    total_amount: 0,
    quote_date: "2025-06-01",
    valid_through: "2025-08-31",
    file_name: "LoneStarConcrete_Quote_Jun2025.pdf",
    file_type: "PDF",
    notes: "Includes mobilization. Excludes subgrade prep.",
    status: "Approved",
    uploaded_by: "J. Rourke",
    uploaded_at: "2025-06-03",
  },
  {
    id: "Q-002",
    org_id: "org-1",
    company: "Texas Steel Erectors",
    trade: "Structural Steel",
    contact_name: "Dana Willcox",
    email: "dwillcox@txsteel.com",
    phone: "972-555-0341",
    market: "Dallas–Fort Worth",
    cost_item_id: "CI-011",
    csi_code: "05 10 00",
    csi_label: "Structural Metal Framing",
    description: "A992 W-Shape, fab & erect, includes shop drawings",
    unit: "TON",
    unit_cost: 24200000,
    total_amount: 0,
    quote_date: "2025-06-05",
    valid_through: "2025-09-05",
    file_name: "TXSteel_BidPackage_Jun2025.xlsx",
    file_type: "XLSX",
    notes: "Lead time 10–12 weeks from award.",
    status: "Reviewed",
    uploaded_by: "J. Rourke",
    uploaded_at: "2025-06-06",
  },
  {
    id: "Q-003",
    org_id: "org-1",
    company: "Gulf Coast Roofing",
    trade: "Roofing",
    contact_name: "Sandra Kim",
    email: "skim@gulfcoastroof.com",
    phone: "713-555-0094",
    market: "Houston",
    cost_item_id: "CI-025",
    csi_code: "07 50 00",
    csi_label: "Membrane Roofing",
    description: "60mil TPO fully adhered, R-25 insulation board",
    unit: "SF",
    unit_cost: 418,
    total_amount: 0,
    quote_date: "2025-06-12",
    valid_through: "2025-09-12",
    file_name: "GulfCoast_TPO_Quote.csv",
    file_type: "CSV",
    notes: "Warranty: 20yr manufacturer + 2yr labor.",
    status: "Received",
    uploaded_by: "M. Chen",
    uploaded_at: "2025-06-13",
  },
];


// ─── Helpers ──────────────────────────────────────────────────────────────────
const cents = (v) => (Number(v) || 0);
const fmtCost = (v) => {
  if (v === null || v === undefined || v === "") return "—";
  const n = Number(v);
  if (isNaN(n)) return "—";
  return "$" + (n / 100).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
};
const fmtDate = (iso) => {
  if (!iso) return "—";
  const d = new Date(iso.includes("T") ? iso : iso + "T00:00:00");
  return d.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });
};
const totalCost = (item) => cents(item.mat_cost) + cents(item.lab_cost) + cents(item.equip_cost);

// ─── Shared Styles ────────────────────────────────────────────────────────────
const s = {
  shell: { display: "flex", height: "100vh", width: "100%", fontFamily: "'Inter', system-ui, -apple-system, sans-serif", backgroundColor: C.contentBg, overflow: "hidden", WebkitFontSmoothing: "antialiased", MozOsxFontSmoothing: "grayscale" },

  // NAV
  nav: { width: 228, minWidth: 228, backgroundColor: C.navBg, display: "flex", flexDirection: "column", flexShrink: 0, borderRight: "1px solid #131C2A" },
  navHeader: { padding: "22px 18px 16px", borderBottom: "1px solid #131C2A" },
  navOrgName: { fontSize: 13, fontWeight: 700, color: C.navActive, lineHeight: 1.3, marginBottom: 3, letterSpacing: "-0.01em" },
  navOrgPlan: { fontSize: 10, color: "#4A5C75", textTransform: "uppercase", letterSpacing: "0.1em", fontWeight: 600 },
  navSection: { padding: "16px 10px 4px" },
  navSectionLabel: { fontSize: 10, fontWeight: 700, color: "#374357", textTransform: "uppercase", letterSpacing: "0.12em", paddingLeft: 8, marginBottom: 6 },
  navItem: (active) => ({ display: "flex", alignItems: "center", gap: 9, padding: "8px 10px", borderRadius: 6, cursor: "pointer", backgroundColor: active ? "#1A2840" : "transparent", color: active ? C.navActive : C.navText, fontSize: 13, fontWeight: active ? 600 : 400, userSelect: "none", transition: "all 0.1s" }),
  navDot: (active) => ({ width: 6, height: 6, borderRadius: "50%", backgroundColor: active ? C.accentLight : "transparent", border: active ? "none" : "1px solid #2D3E54", flexShrink: 0 }),
  navFooter: { marginTop: "auto", padding: "14px 18px", borderTop: "1px solid #131C2A", display: "flex", alignItems: "center", gap: 10 },
  avatar: { width: 32, height: 32, borderRadius: 8, backgroundColor: C.accentLight, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 12, fontWeight: 700, color: "#fff", flexShrink: 0 },
  navUserName: { fontSize: 13, fontWeight: 600, color: C.navActive },
  navUserRole: { fontSize: 11, color: C.navText },

  // MAIN
  main: { flex: 1, display: "flex", flexDirection: "column", overflow: "hidden" },
  topBar: { height: 54, backgroundColor: C.white, borderBottom: `1px solid ${C.border}`, display: "flex", alignItems: "center", padding: "0 24px", gap: 16, flexShrink: 0, boxShadow: "0 1px 3px rgba(0,0,0,0.04)" },
  topBarTitle: { fontSize: 15, fontWeight: 600, color: C.textPrimary, flex: 1 },

  // BREADCRUMB
  breadcrumb: { padding: "10px 24px", backgroundColor: C.white, borderBottom: `1px solid ${C.border}`, display: "flex", alignItems: "center", gap: 6, fontSize: 13, flexShrink: 0 },
  crumbLink: { color: C.accentLight, cursor: "pointer", fontWeight: 500 },
  crumbSep: { color: C.textLight },
  crumbCurrent: { color: C.textPrimary, fontWeight: 500 },

  // TOOLBAR
  toolbar: { display: "flex", alignItems: "center", gap: 10, padding: "10px 24px", borderBottom: `1px solid ${C.border}`, backgroundColor: C.white, flexShrink: 0 },
  searchInput: { height: 34, padding: "0 12px", fontSize: 13, border: `1px solid ${C.border}`, borderRadius: 6, outline: "none", width: 280, color: C.textPrimary, backgroundColor: C.contentBg, fontFamily: "inherit", transition: "border 0.15s" },
  select: { height: 34, padding: "0 10px", fontSize: 13, border: `1px solid ${C.border}`, borderRadius: 6, outline: "none", color: C.textPrimary, backgroundColor: C.contentBg, fontFamily: "inherit", cursor: "pointer" },
  spacer: { flex: 1 },
  btn: (variant = "secondary") => ({ height: 34, padding: "0 16px", fontSize: 13, borderRadius: 6, cursor: "pointer", fontFamily: "inherit", fontWeight: 600, whiteSpace: "nowrap", border: variant === "primary" ? "none" : `1px solid ${C.border}`, backgroundColor: variant === "primary" ? C.accentLight : variant === "danger" ? C.danger : C.white, color: variant === "primary" || variant === "danger" ? "#fff" : C.textPrimary, display: "inline-flex", alignItems: "center", gap: 6, boxShadow: variant === "primary" ? "0 1px 3px rgba(37,99,235,0.3)" : "0 1px 2px rgba(0,0,0,0.05)", transition: "all 0.15s" }),

  // TABLE
  tableWrap: { flex: 1, overflow: "auto" },
  table: { width: "100%", borderCollapse: "collapse", fontSize: 13 },
  thead: { backgroundColor: C.contentBg, position: "sticky", top: 0, zIndex: 1 },
  th: (sortable) => ({ padding: "9px 14px", textAlign: "left", borderBottom: `2px solid ${C.border}`, fontSize: 11, fontWeight: 700, color: C.textMuted, textTransform: "uppercase", letterSpacing: "0.08em", whiteSpace: "nowrap", cursor: sortable ? "pointer" : "default", userSelect: "none", backgroundColor: C.contentBg }),
  thRight: (sortable) => ({ padding: "9px 14px", textAlign: "right", borderBottom: `2px solid ${C.border}`, fontSize: 11, fontWeight: 700, color: C.textMuted, textTransform: "uppercase", letterSpacing: "0.08em", whiteSpace: "nowrap", cursor: sortable ? "pointer" : "default", userSelect: "none", backgroundColor: C.contentBg }),
  td: { padding: "10px 14px", borderBottom: `1px solid ${C.border}`, color: C.textPrimary, whiteSpace: "nowrap" },
  tdRight: { padding: "10px 14px", borderBottom: `1px solid ${C.border}`, color: C.textPrimary, whiteSpace: "nowrap", textAlign: "right", fontFamily: "'DM Mono','Roboto Mono',monospace", fontSize: 12 },
  tdMono: { padding: "10px 14px", borderBottom: `1px solid ${C.border}`, color: C.textMuted, fontFamily: "'DM Mono','Roboto Mono',monospace", fontSize: 12, whiteSpace: "nowrap" },
  tdMuted: { padding: "10px 14px", borderBottom: `1px solid ${C.border}`, color: C.textMuted, whiteSpace: "nowrap" },
  emptyRow: { padding: "56px 24px", textAlign: "center", color: C.textMuted, fontSize: 13 },

  // SOURCE BADGE
  sourceBadge: (src) => {
    const map = { Bid: { bg: "#EFF6FF", color: "#1D4ED8" }, Estimate: { bg: "#F5F3FF", color: "#6D28D9" }, Actuals: { bg: C.successBg, color: C.success }, "RS Means": { bg: C.warningBg, color: C.warning }, Other: { bg: C.neutralBg, color: C.neutral } };
    const t = map[src] || map.Other;
    return { display: "inline-flex", alignItems: "center", padding: "2px 8px", borderRadius: 4, fontSize: 11, fontWeight: 600, backgroundColor: t.bg, color: t.color, textTransform: "uppercase", letterSpacing: "0.05em" };
  },

  // ACCORDION — DIVISION ROWS
  accordionWrap: { flex: 1, overflowY: "auto", backgroundColor: C.contentBg },
  divRow: (open, hover) => ({
    display: "flex", alignItems: "center",
    padding: "0 24px", height: 46,
    backgroundColor: open ? "#EEF4FF" : hover ? "#F0F4F8" : C.white,
    borderBottom: `1px solid ${C.border}`,
    cursor: "pointer", userSelect: "none",
    transition: "background 0.1s",
  }),
  divRowArrow: (open) => ({
    fontSize: 9, color: C.accentLight, marginRight: 12, flexShrink: 0,
    transform: open ? "rotate(90deg)" : "rotate(0deg)",
    transition: "transform 0.15s",
    display: "inline-block",
  }),
  divRowCode: {
    fontSize: 11, fontWeight: 700, color: C.accentLight,
    fontFamily: "'DM Mono','Roboto Mono',monospace",
    letterSpacing: "0.06em", marginRight: 14, flexShrink: 0, width: 30,
  },
  divRowLabel: (open) => ({
    fontSize: 12, fontWeight: 700, color: C.textPrimary,
    textTransform: "uppercase", letterSpacing: "0.07em", flex: 1,
  }),
  divRowCount: {
    fontSize: 12, color: C.textMuted, fontFamily: "'DM Mono','Roboto Mono',monospace",
    flexShrink: 0,
  },

  // ACCORDION — SUBDIVISION ROWS
  subPanel: { backgroundColor: C.contentBg, borderBottom: `1px solid ${C.border}` },
  subAccRow: (hover) => ({
    display: "flex", alignItems: "center",
    padding: "0 24px 0 56px", height: 40,
    backgroundColor: hover ? C.tableHover : C.tableStripe,
    borderBottom: `1px solid ${C.border}`,
    cursor: "pointer", userSelect: "none",
    transition: "background 0.1s",
  }),
  subAccArrow: { fontSize: 9, color: C.textLight, marginRight: 12, flexShrink: 0 },
  subAccCode: {
    fontSize: 11, color: C.textMuted,
    fontFamily: "'DM Mono','Roboto Mono',monospace",
    marginRight: 16, flexShrink: 0, width: 76,
  },
  subAccLabel: { fontSize: 13, fontWeight: 400, color: C.textPrimary, flex: 1 },
  subAccCount: { fontSize: 12, color: C.textMuted, fontFamily: "'DM Mono','Roboto Mono',monospace" },

  // TABLE FOOTER
  tableFooter: { padding: "10px 24px", borderTop: `1px solid ${C.border}`, backgroundColor: C.white, fontSize: 12, color: C.textMuted, flexShrink: 0, display: "flex", alignItems: "center", gap: 16 },

  // MODAL / OVERLAY
  overlay: { position: "fixed", inset: 0, backgroundColor: "rgba(10,18,30,0.55)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 100, backdropFilter: "blur(2px)" },
  modal: (wide) => ({ backgroundColor: C.white, borderRadius: 10, width: wide ? 660 : 540, maxHeight: "92vh", boxShadow: "0 20px 60px rgba(0,0,0,0.18), 0 4px 16px rgba(0,0,0,0.08)", overflow: "hidden", display: "flex", flexDirection: "column" }),
  modalHeader: { padding: "18px 22px", borderBottom: `1px solid ${C.border}`, display: "flex", alignItems: "center", justifyContent: "space-between", flexShrink: 0 },
  modalTitle: { fontSize: 15, fontWeight: 700, color: C.textPrimary, letterSpacing: "-0.01em" },
  modalSubtitle: { fontSize: 12, color: C.textMuted, marginTop: 3 },
  modalBody: { padding: 22, overflowY: "auto", display: "flex", flexDirection: "column", gap: 14 },
  modalFooter: { padding: "14px 22px", borderTop: `1px solid ${C.border}`, display: "flex", justifyContent: "space-between", alignItems: "center", flexShrink: 0, backgroundColor: C.contentBg },
  closeBtn: { background: "none", border: "none", cursor: "pointer", color: C.textMuted, fontSize: 22, lineHeight: 1, padding: "2px 4px", borderRadius: 4 },
  field: { display: "flex", flexDirection: "column", gap: 5 },
  label: { fontSize: 12, fontWeight: 600, color: C.textMuted, letterSpacing: "0.01em" },
  input: { height: 36, padding: "0 11px", fontSize: 13, border: `1px solid ${C.border}`, borderRadius: 6, outline: "none", color: C.textPrimary, backgroundColor: C.white, fontFamily: "inherit", width: "100%", boxSizing: "border-box", transition: "border 0.15s", boxShadow: "0 1px 2px rgba(0,0,0,0.04)" },
  textarea: { padding: "9px 11px", fontSize: 13, border: `1px solid ${C.border}`, borderRadius: 6, outline: "none", color: C.textPrimary, backgroundColor: C.white, fontFamily: "inherit", width: "100%", boxSizing: "border-box", resize: "vertical", minHeight: 64, boxShadow: "0 1px 2px rgba(0,0,0,0.04)" },
  fieldSelect: { height: 36, padding: "0 11px", fontSize: 13, border: `1px solid ${C.border}`, borderRadius: 6, outline: "none", color: C.textPrimary, backgroundColor: C.white, fontFamily: "inherit", width: "100%", boxSizing: "border-box", cursor: "pointer", boxShadow: "0 1px 2px rgba(0,0,0,0.04)" },
  grid2: { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 },
  grid3: { display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 14 },
  sectionDivider: { fontSize: 11, fontWeight: 700, color: C.textMuted, textTransform: "uppercase", letterSpacing: "0.1em", paddingTop: 6, borderBottom: `1px solid ${C.border}`, paddingBottom: 7, marginTop: 2 },
  costTotal: { display: "flex", justifyContent: "flex-end", alignItems: "center", gap: 8, padding: "10px 14px", backgroundColor: C.contentBg, borderRadius: 6, fontSize: 13, border: `1px solid ${C.border}` },
  errorMsg: { fontSize: 12, color: C.danger, background: C.dangerBg, padding: "9px 12px", borderRadius: 6, border: `1px solid #FECACA` },

  // HISTORY
  histEntry: { borderLeft: `3px solid ${C.border}`, paddingLeft: 14, paddingBottom: 16 },
  histDate: { fontSize: 11, color: C.textMuted, marginBottom: 4 },
  histUser: { fontSize: 12, fontWeight: 600, color: C.textPrimary, marginBottom: 6 },
  histRow: { display: "flex", gap: 16, fontSize: 12 },
  histCell: { display: "flex", flexDirection: "column", gap: 2 },
  histCellLabel: { fontSize: 10, color: C.textMuted, textTransform: "uppercase", letterSpacing: "0.06em" },
  histCellValue: { fontSize: 13, fontFamily: "'DM Mono','Roboto Mono',monospace", color: C.textPrimary },
};

// ─── State → Metro Select ─────────────────────────────────────────────────────
// Two-step geography selector: State dropdown → Metro dropdown.
// value / onChange operate on the metro string (matches item.region).
// inline=true renders as two side-by-side selects for compact contexts.
function StateMetroSelect({ value, onChange, inline = false }) {
  // Derive the current state from the current value
  const currentState = Object.keys(STATE_METROS).find(st =>
    STATE_METROS[st].includes(value)
  ) || "";

  const [selectedState, setSelectedState] = useState(currentState);

  // When parent resets value externally, sync local state
  useEffect(() => {
    const st = Object.keys(STATE_METROS).find(s => STATE_METROS[s].includes(value)) || "";
    if (st !== selectedState) setSelectedState(st);
  }, [value]);

  function handleStateChange(e) {
    const st = e.target.value;
    setSelectedState(st);
    onChange(""); // clear metro when state changes
  }

  function handleMetroChange(e) {
    onChange(e.target.value);
  }

  const metros = selectedState ? STATE_METROS[selectedState] : [];

  if (inline) {
    // Compact side-by-side layout for use inside grid rows
    return (
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
        <div style={s.field}>
          <label style={s.label}>State</label>
          <select style={s.fieldSelect} value={selectedState} onChange={handleStateChange}>
            <option value="">— State —</option>
            {ALL_STATES.map(st => <option key={st} value={st}>{st}</option>)}
          </select>
        </div>
        <div style={s.field}>
          <label style={s.label}>Market *</label>
          <select
            style={s.fieldSelect}
            value={value}
            onChange={handleMetroChange}
            disabled={!selectedState}
          >
            <option value="">— Market —</option>
            {metros.map(m => <option key={m} value={m}>{m}</option>)}
          </select>
        </div>
      </div>
    );
  }

  // Stacked layout for modal forms
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
      <div style={s.grid2}>
        <div style={s.field}>
          <label style={s.label}>State</label>
          <select style={s.fieldSelect} value={selectedState} onChange={handleStateChange}>
            <option value="">— Select state —</option>
            {ALL_STATES.map(st => <option key={st} value={st}>{st}</option>)}
          </select>
        </div>
        <div style={s.field}>
          <label style={s.label}>Market / Metro</label>
          <select
            style={s.fieldSelect}
            value={value}
            onChange={handleMetroChange}
            disabled={!selectedState}
          >
            <option value="">{selectedState ? "— Select market —" : "— Select state first —"}</option>
            {metros.map(m => <option key={m} value={m}>{m}</option>)}
          </select>
        </div>
      </div>
    </div>
  );
}

// ─── Sort Icon ────────────────────────────────────────────────────────────────
function SortIcon({ col, sortCol, sortDir }) {
  if (col !== sortCol) return <span style={{ color: C.textLight, marginLeft: 3, fontSize: 10 }}>↕</span>;
  return <span style={{ color: C.accentLight, marginLeft: 3, fontSize: 10 }}>{sortDir === "asc" ? "↑" : "↓"}</span>;
}

// ─── Edit / Add Modal ─────────────────────────────────────────────────────────
function CostItemModal({ item, subdivisionCode, subdivisionLabel, onClose, onSave, onDelete }) {
  const isNew = !item;
  const [form, setForm] = useState(item ? {
    description: item.description,
    unit: item.unit,
    mat_cost: (item.mat_cost / 100).toFixed(2),
    lab_cost: (item.lab_cost / 100).toFixed(2),
    equip_cost: (item.equip_cost / 100).toFixed(2),
    region: item.region,
    source: item.source,
    notes: item.notes || "",
  } : {
    description: "", unit: "SF", mat_cost: "", lab_cost: "", equip_cost: "",
    region: "", source: "Estimate", notes: "",
  });
  const [error, setError] = useState("");
  const [confirmDelete, setConfirmDelete] = useState(false);

  const f = (field) => (e) => setForm(p => ({ ...p, [field]: e.target.value }));

  const parsedTotal = (
    (parseFloat(form.mat_cost)  || 0) +
    (parseFloat(form.lab_cost)  || 0) +
    (parseFloat(form.equip_cost)|| 0)
  );

  function handleSave() {
    if (!form.description.trim()) { setError("Description is required."); return; }
    if (!form.unit) { setError("Unit of measure is required."); return; }
    const saved = {
      ...(item || {}),
      id: item?.id || generateId(),
      org_id: "org-1",
      csi_code: subdivisionCode,
      description: form.description.trim(),
      unit: form.unit,
      mat_cost:   Math.round((parseFloat(form.mat_cost)  || 0) * 100),
      lab_cost:   Math.round((parseFloat(form.lab_cost)  || 0) * 100),
      equip_cost: Math.round((parseFloat(form.equip_cost)|| 0) * 100),
      region: form.region,
      source: form.source,
      notes: form.notes.trim(),
      updated_by: MOCK_USER.name.split(" ").map(n => n[0]).join(". ") + ".",
      updated_at: new Date().toISOString().slice(0, 10),
    };
    onSave(saved, !isNew);
    onClose();
  }

  return (
    <div style={s.overlay} onClick={(e) => e.target === e.currentTarget && onClose()}>
      <div style={s.modal(true)}>
        <div style={s.modalHeader}>
          <div>
            <div style={s.modalTitle}>{isNew ? "Add Cost Item" : "Edit Cost Item"}</div>
            <div style={s.modalSubtitle}>{subdivisionCode} — {subdivisionLabel}</div>
          </div>
          <button style={s.closeBtn} onClick={onClose}>×</button>
        </div>

        <div style={s.modalBody}>
          {error && <div style={s.errorMsg}>{error}</div>}

          <div style={s.field}>
            <label style={s.label}>Description *</label>
            <input style={s.input} value={form.description} onChange={f("description")} placeholder='e.g. Slab on Grade – 4" w/ WWF' autoFocus />
          </div>

          <div style={s.grid2}>
            <div style={s.field}>
              <label style={s.label}>Unit of Measure *</label>
              <select style={s.fieldSelect} value={form.unit} onChange={f("unit")}>
                {UNITS.map(u => <option key={u} value={u}>{u}</option>)}
              </select>
            </div>
            <div style={s.field}>
              <label style={s.label}>Source</label>
              <select style={s.fieldSelect} value={form.source} onChange={f("source")}>
                {SOURCES.map(src => <option key={src} value={src}>{src}</option>)}
              </select>
            </div>
          </div>

          <div style={s.sectionDivider}>Unit Costs</div>

          <div style={s.grid3}>
            {[["mat_cost","Material"],["lab_cost","Labor"],["equip_cost","Equipment"]].map(([key, lbl]) => (
              <div key={key} style={s.field}>
                <label style={s.label}>{lbl} ($/unit)</label>
                <input
                  style={{ ...s.input, fontFamily: "'DM Mono','Roboto Mono',monospace" }}
                  type="number" min="0" step="0.01"
                  value={form[key]} onChange={f(key)}
                  placeholder="0.00"
                />
              </div>
            ))}
          </div>

          <div style={s.costTotal}>
            <span style={{ color: C.textMuted, fontSize: 12 }}>Total Unit Cost:</span>
            <span style={{ fontWeight: 700, fontSize: 15, fontFamily: "'DM Mono','Roboto Mono',monospace", color: C.textPrimary }}>
              ${parsedTotal.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
            </span>
          </div>

          <div style={s.sectionDivider}>Context</div>
          <StateMetroSelect
            value={form.region}
            onChange={region => setForm(p => ({ ...p, region }))}
          />

          <div style={s.field}>
            <label style={s.label}>Notes / Crew & Productivity</label>
            <textarea style={s.textarea} value={form.notes} onChange={f("notes")} placeholder="Scope inclusions, crew size, lead time notes…" />
          </div>
        </div>

        <div style={s.modalFooter}>
          <div>
            {!isNew && !confirmDelete && (
              <button style={s.btn("danger")} onClick={() => setConfirmDelete(true)}>Delete</button>
            )}
            {confirmDelete && (
              <span style={{ fontSize: 12, color: C.danger, display: "flex", alignItems: "center", gap: 8 }}>
                Confirm delete?
                <button style={s.btn("danger")} onClick={() => { onDelete(item.id); onClose(); }}>Yes, delete</button>
                <button style={s.btn()} onClick={() => setConfirmDelete(false)}>Cancel</button>
              </span>
            )}
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            <button style={s.btn()} onClick={onClose}>Cancel</button>
            <button style={s.btn("primary")} onClick={handleSave}>{isNew ? "Add Cost Item" : "Save Changes"}</button>
          </div>
        </div>
      </div>
    </div>
  );
}

// ─── Item Detail Modal (Cost History + Sub Quotes tabs) ──────────────────────
function HistoryModal({ item, onClose, quotes, onUpdateQuotes }) {
  const [activeTab, setActiveTab] = useState("history");
  const [viewingQuote, setViewingQuote] = useState(null);
  const history = MOCK_HISTORY[item.id] || [];
  const itemQuotes = (quotes || []).filter(q => q.cost_item_id === item.id);

  const statusColors = {
    Received: { bg: "#EFF6FF", color: "#1D4ED8" },
    Reviewed: { bg: "#F5F3FF", color: "#6D28D9" },
    Approved: { bg: "#DCFCE7", color: "#16A34A" },
    Rejected: { bg: "#FEE2E2", color: "#DC2626" },
  };

  const TAB = ({ id, label, count }) => (
    <button
      onClick={() => setActiveTab(id)}
      style={{
        padding: "8px 16px", fontSize: 13, fontWeight: 500,
        border: "none", cursor: "pointer", background: "none",
        fontFamily: "inherit",
        color: activeTab === id ? C.accentLight : C.textMuted,
        borderBottom: activeTab === id ? `2px solid ${C.accentLight}` : "2px solid transparent",
        marginBottom: -1,
      }}
    >
      {label}
      {count > 0 && (
        <span style={{
          marginLeft: 6, padding: "1px 6px", borderRadius: 10,
          fontSize: 10, fontWeight: 700,
          backgroundColor: activeTab === id ? C.accentLight : C.neutralBg,
          color: activeTab === id ? "#fff" : C.textMuted,
        }}>{count}</span>
      )}
    </button>
  );

  return (
    <>
    <div style={s.overlay} onClick={(e) => e.target === e.currentTarget && onClose()}>
      <div style={{ ...s.modal(true), width: 620 }}>
        {/* Header */}
        <div style={s.modalHeader}>
          <div>
            <div style={s.modalTitle}>{item.id}</div>
            <div style={s.modalSubtitle}>{item.description}</div>
          </div>
          <button style={s.closeBtn} onClick={onClose}>×</button>
        </div>

        {/* Tabs */}
        <div style={{
          display: "flex", borderBottom: `1px solid ${C.border}`,
          padding: "0 20px", backgroundColor: C.white, flexShrink: 0,
        }}>
          <TAB id="history" label="Cost History" count={history.length} />
          <TAB id="quotes"  label="Sub Quotes"   count={itemQuotes.length} />
        </div>

        {/* Tab content */}
        <div style={{ ...s.modalBody, minHeight: 200 }}>

          {/* ── Cost History tab ── */}
          {activeTab === "history" && (
            history.length === 0 ? (
              <div style={{ color: C.textMuted, fontSize: 13, textAlign: "center", padding: "32px 0" }}>
                No history recorded for this item yet.
              </div>
            ) : history.map((h, i) => (
              <div key={i} style={{ ...s.histEntry, borderColor: i === 0 ? C.accentLight : C.border }}>
                <div style={s.histDate}>{fmtDate(h.changed_at)} {new Date(h.changed_at).toLocaleTimeString("en-US", { hour: "2-digit", minute: "2-digit" })}</div>
                <div style={s.histUser}>{h.changed_by} {i === 0 && <span style={{ ...s.sourceBadge(h.snapshot.source), marginLeft: 6 }}>Current</span>}</div>
                <div style={s.histRow}>
                  {[["Material", h.snapshot.mat_cost],["Labor", h.snapshot.lab_cost],["Equipment", h.snapshot.equip_cost]].map(([lbl, v]) => (
                    <div key={lbl} style={s.histCell}>
                      <span style={s.histCellLabel}>{lbl}</span>
                      <span style={s.histCellValue}>{fmtCost(v)}</span>
                    </div>
                  ))}
                  <div style={s.histCell}>
                    <span style={s.histCellLabel}>Total</span>
                    <span style={{ ...s.histCellValue, fontWeight: 700 }}>
                      {fmtCost(cents(h.snapshot.mat_cost) + cents(h.snapshot.lab_cost) + cents(h.snapshot.equip_cost))}
                    </span>
                  </div>
                  <div style={s.histCell}>
                    <span style={s.histCellLabel}>Source</span>
                    <span style={s.sourceBadge(h.snapshot.source)}>{h.snapshot.source}</span>
                  </div>
                </div>
                {h.snapshot.notes && (
                  <div style={{ marginTop: 6, fontSize: 12, color: C.textMuted, fontStyle: "italic" }}>"{h.snapshot.notes}"</div>
                )}
              </div>
            ))
          )}

          {/* ── Sub Quotes tab ── */}
          {activeTab === "quotes" && (
            itemQuotes.length === 0 ? (
              <div style={{ color: C.textMuted, fontSize: 13, textAlign: "center", padding: "32px 0" }}>
                No quotes linked to this line item yet.
                <div style={{ fontSize: 12, color: C.textLight, marginTop: 6 }}>
                  Upload a quote and link it to <strong>{item.id}</strong> using the upload panel above.
                </div>
              </div>
            ) : itemQuotes.map((q, i) => (
              <div key={q.id} style={{
                border: `1px solid ${C.border}`, borderRadius: 4,
                padding: "14px 16px", marginBottom: 10,
                backgroundColor: i % 2 === 0 ? C.white : C.tableStripe,
              }}>
                {/* Quote header row */}
                <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", marginBottom: 10 }}>
                  <div>
                    <div style={{ fontSize: 13, fontWeight: 600, color: C.textPrimary }}>{q.company}</div>
                    <div style={{ fontSize: 11, color: C.textMuted, marginTop: 2 }}>
                      {q.contact_name}{q.contact_name && q.email ? " · " : ""}{q.email}
                    </div>
                  </div>
                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    <span style={{
                      padding: "2px 8px", borderRadius: 3, fontSize: 10, fontWeight: 700,
                      textTransform: "uppercase", letterSpacing: "0.05em",
                      backgroundColor: (statusColors[q.status] || statusColors.Received).bg,
                      color: (statusColors[q.status] || statusColors.Received).color,
                    }}>{q.status}</span>
                    <select
                      value={q.status}
                      onChange={e => {
                        const updated = quotes.map(x => x.id === q.id ? { ...x, status: e.target.value } : x);
                        onUpdateQuotes(updated);
                      }}
                      style={{
                        fontSize: 11, padding: "2px 6px", border: `1px solid ${C.borderDark}`,
                        borderRadius: 3, color: C.textMuted, backgroundColor: C.white,
                        fontFamily: "inherit", cursor: "pointer",
                      }}
                    >
                      {QUOTE_STATUSES.map(st => <option key={st} value={st}>{st}</option>)}
                    </select>
                  </div>
                </div>

                {/* Cost row */}
                <div style={{ display: "flex", gap: 24, marginBottom: 10 }}>
                  {[
                    ["Description", q.description || "—"],
                    ["Unit", q.unit],
                    ["Unit Cost", "$" + (q.unit_cost / 100).toLocaleString("en-US", { minimumFractionDigits: 2 })],
                    ["Market", q.market],
                    ["Quote Date", q.quote_date],
                    ["Valid Through", q.valid_through || "—"],
                  ].map(([label, val]) => (
                    <div key={label} style={{ display: "flex", flexDirection: "column", gap: 2 }}>
                      <span style={{ fontSize: 10, fontWeight: 600, color: C.textMuted, textTransform: "uppercase", letterSpacing: "0.06em" }}>{label}</span>
                      <span style={{ fontSize: 12, color: C.textPrimary, fontWeight: label === "Unit Cost" ? 700 : 400 }}>{val}</span>
                    </div>
                  ))}
                </div>

                {/* Notes */}
                {q.notes && (
                  <div style={{ fontSize: 11, color: C.textMuted, fontStyle: "italic", marginBottom: 8 }}>
                    {q.notes}
                  </div>
                )}

                {/* File link */}
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <span style={{
                    padding: "1px 6px", borderRadius: 2, fontSize: 10, fontWeight: 600,
                    border: `1px solid ${C.borderDark}`, color: C.textMuted,
                  }}>{q.file_type}</span>
                  <span
                    style={{ fontSize: 11, color: C.accentLight, cursor: "pointer", textDecoration: "underline" }}
                    onClick={() => setViewingQuote(q)}
                  >
                    {q.file_name}
                  </span>
                  <span style={{ fontSize: 11, color: C.textLight }}>· Uploaded {q.uploaded_at} by {q.uploaded_by}</span>
                </div>
              </div>
            ))
          )}
        </div>

        <div style={{ ...s.modalFooter, justifyContent: "flex-end" }}>
          <button style={s.btn()} onClick={onClose}>Close</button>
        </div>
      </div>
    </div>

    {viewingQuote && (
      <PDFViewerModal quote={viewingQuote} onClose={() => setViewingQuote(null)} />
    )}
    </>
  );
}

// ─── Line Items Table ─────────────────────────────────────────────────────────
function LineItemsTable({ items, subdivision, onAdd, onEdit, onDelete, quotes, onUpdateQuotes }) {
  const [search, setSearch]   = useState("");
  const [srcFilter, setSrc]   = useState("");
  const [sortCol, setSortCol] = useState("description");
  const [sortDir, setSortDir] = useState("asc");
  const [hoverRow, setHoverRow] = useState(null);
  const [historyItem, setHistoryItem] = useState(null);

  function toggleSort(col) {
    if (col === sortCol) setSortDir(d => d === "asc" ? "desc" : "asc");
    else { setSortCol(col); setSortDir("asc"); }
  }

  const filtered = useMemo(() => {
    let rows = items.filter(i => i.csi_code === subdivision.code);
    if (search) {
      const q = search.toLowerCase();
      rows = rows.filter(r => r.description.toLowerCase().includes(q) || r.id.toLowerCase().includes(q));
    }
    if (srcFilter) rows = rows.filter(r => r.source === srcFilter);
    rows.sort((a, b) => {
      let av, bv;
      if (sortCol === "total") { av = totalCost(a); bv = totalCost(b); }
      else { av = a[sortCol] ?? ""; bv = b[sortCol] ?? ""; }
      if (typeof av === "number") return sortDir === "asc" ? av - bv : bv - av;
      return sortDir === "asc" ? String(av).localeCompare(String(bv)) : String(bv).localeCompare(String(av));
    });
    return rows;
  }, [items, subdivision.code, search, srcFilter, sortCol, sortDir]);

  const TH = ({ col, label, right }) => {
    const style = right ? s.thRight(true) : s.th(true);
    return <th style={style} onClick={() => toggleSort(col)}>{label}<SortIcon col={col} sortCol={sortCol} sortDir={sortDir} /></th>;
  };

  return (
    <>
      <div style={s.toolbar}>
        <input style={s.searchInput} placeholder="Search line items…" value={search} onChange={e => setSearch(e.target.value)} />
        <select style={s.select} value={srcFilter} onChange={e => setSrc(e.target.value)}>
          <option value="">All Sources</option>
          {SOURCES.map(src => <option key={src} value={src}>{src}</option>)}
        </select>
        <div style={s.spacer} />
        <button style={s.btn("primary")} onClick={() => onAdd(subdivision)}>+ Add Cost Item</button>
      </div>

      <div style={s.tableWrap}>
        <table style={s.table}>
          <thead style={s.thead}>
            <tr>
              <th style={s.th(false)} onClick={() => toggleSort("id")}>ID<SortIcon col="id" sortCol={sortCol} sortDir={sortDir} /></th>
              <TH col="description" label="Description" />
              <th style={s.th(false)}>Unit</th>
              <TH col="mat_cost"  label="Material"  right />
              <TH col="lab_cost"  label="Labor"     right />
              <TH col="equip_cost" label="Equipment" right />
              <TH col="total"     label="Total / Unit" right />
              <TH col="source"    label="Source" />
              <TH col="region"    label="Market" />
              <TH col="updated_at" label="Updated" />
              <th style={s.th(false)}></th>
            </tr>
          </thead>
          <tbody>
            {filtered.length === 0 ? (
              <tr><td colSpan={11} style={s.emptyRow}>
                No items found.{" "}
                {(search || srcFilter) && <span style={{ color: C.accentLight, cursor: "pointer" }} onClick={() => { setSearch(""); setSrc(""); }}>Clear filters</span>}
              </td></tr>
            ) : filtered.map((item, i) => (
              <tr
                key={item.id}
                style={{ backgroundColor: hoverRow === item.id ? C.tableHover : i % 2 === 1 ? C.tableStripe : C.white, cursor: "pointer" }}
                onMouseEnter={() => setHoverRow(item.id)}
                onMouseLeave={() => setHoverRow(null)}
                onClick={() => setHistoryItem(item)}
              >
                <td style={s.tdMono}>{item.id}</td>
                <td style={{ ...s.td, fontWeight: 500, maxWidth: 300, overflow: "hidden", textOverflow: "ellipsis" }}>{item.description}</td>
                <td style={s.tdMono}>{item.unit}</td>
                <td style={s.tdRight}>{fmtCost(item.mat_cost)}</td>
                <td style={s.tdRight}>{fmtCost(item.lab_cost)}</td>
                <td style={s.tdRight}>{fmtCost(item.equip_cost)}</td>
                <td style={{ ...s.tdRight, fontWeight: 700, color: C.textPrimary }}>{fmtCost(totalCost(item))}</td>
                <td style={s.td}><span style={s.sourceBadge(item.source)}>{item.source}</span></td>
                <td style={s.tdMuted}>{item.region}</td>
                <td style={s.tdMuted}>{fmtDate(item.updated_at)}</td>
                <td style={{ ...s.td, padding: "9px 8px", whiteSpace: "nowrap" }}>
                  {(() => {
                    const qCount = (quotes || []).filter(q => q.cost_item_id === item.id).length;
                    return qCount > 0 ? (
                      <span style={{
                        display: "inline-flex", alignItems: "center", gap: 4,
                        marginRight: 10, padding: "1px 7px", borderRadius: 3,
                        backgroundColor: "#EFF6FF", color: C.accentLight,
                        fontSize: 10, fontWeight: 700, cursor: "pointer",
                      }} onClick={e => { e.stopPropagation(); setHistoryItem(item); }}>
                        📄 {qCount}
                      </span>
                    ) : null;
                  })()}
                  <span style={{ color: C.accentLight, fontSize: 12, fontWeight: 500, cursor: "pointer" }}
                    onClick={e => { e.stopPropagation(); onEdit(item, subdivision); }}>Edit</span>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div style={s.tableFooter}>
        <span>{filtered.length} of {items.filter(i => i.csi_code === subdivision.code).length} items</span>
        <span style={{ color: C.textLight }}>Click a row to view cost history & sub quotes</span>
      </div>

      {historyItem && <HistoryModal item={historyItem} onClose={() => setHistoryItem(null)} quotes={quotes} onUpdateQuotes={onUpdateQuotes} />}
    </>
  );
}


// ─── State → Metro Mapping ────────────────────────────────────────────────────
// Ordered by size/relevance within each state — first entry is primary metro.
const STATE_METROS = {
  "Texas":          ["Dallas–Fort Worth", "Houston", "Austin", "San Antonio", "El Paso", "Lubbock", "Amarillo"],
  "Arizona":        ["Phoenix", "Tucson", "Flagstaff"],
  "Colorado":       ["Denver", "Colorado Springs", "Fort Collins", "Grand Junction"],
  "Georgia":        ["Atlanta", "Savannah", "Augusta", "Columbus"],
  "Tennessee":      ["Nashville", "Memphis", "Knoxville", "Chattanooga"],
  "Florida":        ["Miami", "Orlando", "Tampa", "Jacksonville", "Fort Lauderdale", "Tallahassee"],
  "California":     ["Los Angeles", "San Francisco", "San Diego", "Sacramento", "Fresno", "Bakersfield"],
  "Nevada":         ["Las Vegas", "Reno"],
  "North Carolina": ["Charlotte", "Raleigh", "Greensboro", "Winston-Salem", "Wilmington"],
  "Virginia":       ["Northern Virginia", "Richmond", "Virginia Beach", "Roanoke"],
  "Washington":     ["Seattle", "Spokane", "Tacoma", "Bellevue"],
  "Illinois":       ["Chicago", "Springfield", "Rockford", "Peoria"],
  "Ohio":           ["Columbus", "Cleveland", "Cincinnati", "Dayton", "Akron"],
  "Michigan":       ["Detroit", "Grand Rapids", "Lansing", "Ann Arbor"],
  "Minnesota":      ["Minneapolis", "St. Paul", "Rochester", "Duluth"],
  "Missouri":       ["Kansas City", "St. Louis", "Springfield", "Columbia"],
  "Oklahoma":       ["Oklahoma City", "Tulsa", "Norman", "Lawton"],
  "Arkansas":       ["Little Rock", "NWA (Fayetteville–Springdale–Rogers)", "Fort Smith", "Jonesboro"],
  "Louisiana":      ["New Orleans", "Baton Rouge", "Shreveport", "Lafayette"],
  "Mississippi":    ["Jackson", "Gulfport–Biloxi", "Hattiesburg"],
  "Alabama":        ["Birmingham", "Huntsville", "Mobile", "Montgomery"],
  "South Carolina": ["Charleston", "Greenville", "Columbia", "Myrtle Beach"],
  "Kentucky":       ["Louisville", "Lexington", "Bowling Green", "Covington"],
  "Indiana":        ["Indianapolis", "Fort Wayne", "Evansville", "South Bend"],
  "Wisconsin":      ["Milwaukee", "Madison", "Green Bay", "Racine"],
  "Maryland":       ["Baltimore", "Frederick", "Annapolis"],
  "Pennsylvania":   ["Philadelphia", "Pittsburgh", "Allentown", "Harrisburg", "Scranton"],
  "New York":       ["New York City", "Buffalo", "Rochester", "Albany", "Syracuse"],
  "Massachusetts":  ["Boston", "Worcester", "Springfield"],
  "New Jersey":     ["Newark", "Jersey City", "Trenton", "Camden"],
  "Connecticut":    ["Hartford", "New Haven", "Bridgeport", "Stamford"],
  "Oregon":         ["Portland", "Eugene", "Salem", "Bend"],
  "Utah":           ["Salt Lake City", "Provo", "Ogden", "St. George"],
  "New Mexico":     ["Albuquerque", "Santa Fe", "Las Cruces"],
  "Idaho":          ["Boise", "Nampa", "Idaho Falls", "Pocatello"],
  "Montana":        ["Billings", "Missoula", "Great Falls", "Bozeman"],
  "Wyoming":        ["Cheyenne", "Casper", "Laramie"],
  "Nebraska":       ["Omaha", "Lincoln", "Grand Island"],
  "Kansas":         ["Wichita", "Kansas City (KS)", "Topeka"],
  "Iowa":           ["Des Moines", "Cedar Rapids", "Davenport", "Sioux City"],
  "South Dakota":   ["Sioux Falls", "Rapid City"],
  "North Dakota":   ["Fargo", "Bismarck", "Grand Forks"],
  "Alaska":         ["Anchorage", "Fairbanks", "Juneau"],
  "Hawaii":         ["Honolulu", "Maui", "Hilo"],
  "Maine":          ["Portland", "Bangor", "Augusta"],
  "New Hampshire":  ["Manchester", "Nashua", "Concord"],
  "Vermont":        ["Burlington", "Montpelier", "Rutland"],
  "Rhode Island":   ["Providence", "Warwick", "Cranston"],
  "Delaware":       ["Wilmington", "Dover", "Newark"],
  "West Virginia":  ["Charleston", "Huntington", "Morgantown"],
};

// MARKETS derived from STATE_METROS — full list of all metros across all states
const MARKETS = Object.values(STATE_METROS).flat();
const ALL_STATES = Object.keys(STATE_METROS).sort();

// ─── Metro Fallback Chains ────────────────────────────────────────────────────
// Intra-state alternatives listed first, then nearest out-of-state metros.
const METRO_FALLBACK = {
  "Dallas–Fort Worth":                    ["Houston", "Austin", "San Antonio", "Oklahoma City", "Shreveport"],
  "Houston":                              ["Dallas–Fort Worth", "Austin", "San Antonio", "New Orleans", "Baton Rouge"],
  "Austin":                               ["San Antonio", "Dallas–Fort Worth", "Houston"],
  "San Antonio":                          ["Austin", "Dallas–Fort Worth", "Houston"],
  "El Paso":                              ["Albuquerque", "Dallas–Fort Worth", "Tucson"],
  "Lubbock":                              ["Amarillo", "Dallas–Fort Worth", "Albuquerque"],
  "Amarillo":                             ["Lubbock", "Oklahoma City", "Dallas–Fort Worth"],
  "Little Rock":                          ["NWA (Fayetteville–Springdale–Rogers)", "Fort Smith", "Jonesboro", "Dallas–Fort Worth", "Memphis"],
  "NWA (Fayetteville–Springdale–Rogers)": ["Fort Smith", "Little Rock", "Jonesboro", "Tulsa", "Dallas–Fort Worth"],
  "Fort Smith":                           ["NWA (Fayetteville–Springdale–Rogers)", "Little Rock", "Oklahoma City", "Tulsa"],
  "Jonesboro":                            ["Little Rock", "Memphis", "NWA (Fayetteville–Springdale–Rogers)"],
  "Oklahoma City":                        ["Tulsa", "Norman", "Wichita", "Dallas–Fort Worth"],
  "Tulsa":                                ["Oklahoma City", "Norman", "Fort Smith", "Wichita"],
  "Norman":                               ["Oklahoma City", "Tulsa", "Lawton"],
  "Lawton":                               ["Oklahoma City", "Wichita", "Dallas–Fort Worth"],
  "Nashville":                            ["Memphis", "Knoxville", "Chattanooga", "Louisville", "Birmingham"],
  "Memphis":                              ["Nashville", "Jackson", "Little Rock", "Birmingham"],
  "Knoxville":                            ["Chattanooga", "Nashville", "Charlotte", "Atlanta"],
  "Chattanooga":                          ["Knoxville", "Nashville", "Atlanta", "Birmingham"],
  "Birmingham":                           ["Huntsville", "Montgomery", "Mobile", "Nashville", "Atlanta"],
  "Huntsville":                           ["Birmingham", "Nashville", "Chattanooga"],
  "Mobile":                               ["Birmingham", "New Orleans", "Gulfport–Biloxi"],
  "Montgomery":                           ["Birmingham", "Atlanta", "Columbus"],
  "Jackson":                              ["Gulfport–Biloxi", "Hattiesburg", "Memphis", "New Orleans", "Birmingham"],
  "Gulfport–Biloxi":                      ["Jackson", "Hattiesburg", "New Orleans", "Mobile"],
  "Hattiesburg":                          ["Jackson", "Gulfport–Biloxi", "New Orleans"],
  "New Orleans":                          ["Baton Rouge", "Shreveport", "Lafayette", "Houston", "Gulfport–Biloxi"],
  "Baton Rouge":                          ["New Orleans", "Lafayette", "Shreveport", "Houston"],
  "Shreveport":                           ["Baton Rouge", "New Orleans", "Dallas–Fort Worth", "Little Rock"],
  "Lafayette":                            ["Baton Rouge", "New Orleans", "Houston"],
  "Atlanta":                              ["Savannah", "Charlotte", "Birmingham", "Nashville"],
  "Savannah":                             ["Atlanta", "Charleston", "Augusta"],
  "Augusta":                              ["Atlanta", "Columbia", "Charlotte"],
  "Columbus":                             ["Atlanta", "Birmingham", "Montgomery"],
  "Charleston":                           ["Columbia", "Greenville", "Myrtle Beach", "Savannah", "Atlanta"],
  "Greenville":                           ["Columbia", "Charlotte", "Knoxville", "Atlanta"],
  "Columbia":                             ["Charlotte", "Charleston", "Greenville"],
  "Myrtle Beach":                         ["Charleston", "Wilmington", "Charlotte"],
  "Charlotte":                            ["Raleigh", "Greensboro", "Atlanta", "Greenville"],
  "Raleigh":                              ["Charlotte", "Greensboro", "Richmond", "Virginia Beach"],
  "Greensboro":                           ["Charlotte", "Raleigh", "Winston-Salem"],
  "Winston-Salem":                        ["Greensboro", "Charlotte", "Raleigh"],
  "Wilmington":                           ["Raleigh", "Myrtle Beach", "Charlotte"],
  "Northern Virginia":                    ["Richmond", "Baltimore", "Roanoke"],
  "Richmond":                             ["Northern Virginia", "Charlotte", "Roanoke"],
  "Virginia Beach":                       ["Richmond", "Raleigh", "Northern Virginia"],
  "Roanoke":                              ["Richmond", "Charlotte", "Knoxville"],
  "Denver":                               ["Colorado Springs", "Fort Collins", "Salt Lake City", "Albuquerque"],
  "Colorado Springs":                     ["Denver", "Albuquerque"],
  "Fort Collins":                         ["Denver", "Cheyenne", "Colorado Springs"],
  "Grand Junction":                       ["Denver", "Salt Lake City", "Albuquerque"],
  "Phoenix":                              ["Tucson", "Flagstaff", "Albuquerque", "Las Vegas"],
  "Tucson":                               ["Phoenix", "Flagstaff", "El Paso", "Albuquerque"],
  "Flagstaff":                            ["Phoenix", "Tucson", "Albuquerque"],
  "Las Vegas":                            ["Reno", "Phoenix", "Los Angeles", "Salt Lake City"],
  "Reno":                                 ["Las Vegas", "Sacramento", "Salt Lake City"],
  "Los Angeles":                          ["San Diego", "San Francisco", "Las Vegas"],
  "San Francisco":                        ["Sacramento", "Los Angeles", "Portland"],
  "San Diego":                            ["Los Angeles", "Phoenix"],
  "Sacramento":                           ["San Francisco", "Los Angeles", "Portland"],
  "Fresno":                               ["Sacramento", "Los Angeles", "Bakersfield"],
  "Bakersfield":                          ["Fresno", "Los Angeles", "Las Vegas"],
  "Seattle":                              ["Tacoma", "Bellevue", "Spokane", "Portland"],
  "Spokane":                              ["Seattle", "Boise", "Portland"],
  "Tacoma":                               ["Seattle", "Portland", "Bellevue"],
  "Bellevue":                             ["Seattle", "Tacoma", "Portland"],
  "Portland":                             ["Salem", "Eugene", "Seattle", "Boise"],
  "Eugene":                               ["Salem", "Portland", "Bend"],
  "Salem":                                ["Portland", "Eugene", "Seattle"],
  "Bend":                                 ["Portland", "Boise", "Eugene"],
  "Boise":                                ["Nampa", "Salt Lake City", "Portland", "Spokane"],
  "Nampa":                                ["Boise", "Salt Lake City"],
  "Idaho Falls":                          ["Boise", "Salt Lake City", "Pocatello"],
  "Pocatello":                            ["Idaho Falls", "Boise", "Salt Lake City"],
  "Billings":                             ["Missoula", "Bozeman", "Cheyenne", "Denver"],
  "Missoula":                             ["Billings", "Boise", "Spokane"],
  "Great Falls":                          ["Billings", "Missoula", "Bozeman"],
  "Bozeman":                              ["Billings", "Missoula", "Salt Lake City"],
  "Cheyenne":                             ["Casper", "Denver", "Fort Collins"],
  "Casper":                               ["Cheyenne", "Billings", "Denver"],
  "Laramie":                              ["Cheyenne", "Denver", "Fort Collins"],
  "Salt Lake City":                       ["Provo", "Ogden", "Denver", "Las Vegas"],
  "Provo":                                ["Salt Lake City", "Ogden", "Denver"],
  "Ogden":                                ["Salt Lake City", "Provo", "Boise"],
  "St. George":                           ["Salt Lake City", "Las Vegas", "Phoenix"],
  "Albuquerque":                          ["Santa Fe", "Las Cruces", "El Paso", "Phoenix", "Denver"],
  "Santa Fe":                             ["Albuquerque", "Denver", "Phoenix"],
  "Las Cruces":                           ["Albuquerque", "El Paso", "Phoenix"],
  "Chicago":                              ["Milwaukee", "Indianapolis", "Detroit", "St. Louis"],
  "Milwaukee":                            ["Chicago", "Madison", "Minneapolis"],
  "Madison":                              ["Milwaukee", "Chicago", "Minneapolis"],
  "Green Bay":                            ["Milwaukee", "Minneapolis", "Chicago"],
  "Racine":                               ["Milwaukee", "Chicago"],
  "Minneapolis":                          ["St. Paul", "Milwaukee", "Des Moines", "Fargo"],
  "St. Paul":                             ["Minneapolis", "Milwaukee", "Des Moines"],
  "Rochester":                            ["Minneapolis", "Milwaukee", "Duluth"],
  "Duluth":                               ["Minneapolis", "Milwaukee", "Fargo"],
  "Des Moines":                           ["Cedar Rapids", "Kansas City", "Minneapolis", "Omaha"],
  "Cedar Rapids":                         ["Des Moines", "Davenport", "Chicago"],
  "Davenport":                            ["Cedar Rapids", "Des Moines", "Chicago"],
  "Sioux City":                           ["Des Moines", "Omaha", "Sioux Falls"],
  "Sioux Falls":                          ["Rapid City", "Omaha", "Minneapolis"],
  "Rapid City":                           ["Sioux Falls", "Billings", "Denver"],
  "Fargo":                                ["Bismarck", "Minneapolis", "Sioux Falls"],
  "Bismarck":                             ["Fargo", "Sioux Falls", "Minneapolis"],
  "Grand Forks":                          ["Fargo", "Minneapolis"],
  "Omaha":                                ["Lincoln", "Kansas City", "Des Moines"],
  "Lincoln":                              ["Omaha", "Kansas City", "Des Moines"],
  "Grand Island":                         ["Omaha", "Lincoln", "Wichita"],
  "Wichita":                              ["Kansas City (KS)", "Topeka", "Oklahoma City", "Tulsa"],
  "Kansas City (KS)":                     ["Wichita", "Kansas City", "Topeka"],
  "Topeka":                               ["Wichita", "Kansas City", "Omaha"],
  "Kansas City":                          ["St. Louis", "Wichita", "Omaha", "Des Moines"],
  "St. Louis":                            ["Kansas City", "Springfield", "Nashville", "Indianapolis"],
  "Springfield":                          ["Kansas City", "St. Louis", "Columbia"],
  "Columbia":                             ["Kansas City", "St. Louis", "Springfield"],
  "Indianapolis":                         ["Fort Wayne", "Cincinnati", "Chicago", "Louisville"],
  "Fort Wayne":                           ["Indianapolis", "Chicago", "Columbus"],
  "Evansville":                           ["Indianapolis", "Nashville", "Louisville"],
  "South Bend":                           ["Indianapolis", "Chicago", "Detroit"],
  "Columbus":                             ["Cleveland", "Cincinnati", "Indianapolis", "Pittsburgh"],
  "Cleveland":                            ["Columbus", "Pittsburgh", "Detroit", "Akron"],
  "Cincinnati":                           ["Dayton", "Louisville", "Columbus", "Indianapolis"],
  "Dayton":                               ["Cincinnati", "Columbus", "Indianapolis"],
  "Akron":                                ["Cleveland", "Columbus", "Pittsburgh"],
  "Detroit":                              ["Grand Rapids", "Cleveland", "Chicago", "Columbus"],
  "Grand Rapids":                         ["Detroit", "Lansing", "Chicago"],
  "Lansing":                              ["Detroit", "Grand Rapids", "Indianapolis"],
  "Ann Arbor":                            ["Detroit", "Lansing"],
  "Louisville":                           ["Lexington", "Indianapolis", "Nashville", "Cincinnati"],
  "Lexington":                            ["Louisville", "Cincinnati", "Knoxville"],
  "Bowling Green":                        ["Nashville", "Louisville", "Lexington"],
  "Covington":                            ["Louisville", "Cincinnati", "Lexington"],
  "Pittsburgh":                           ["Philadelphia", "Cleveland", "Columbus", "Baltimore"],
  "Philadelphia":                         ["Pittsburgh", "Baltimore", "Newark", "Wilmington"],
  "Allentown":                            ["Philadelphia", "Newark", "New York City"],
  "Harrisburg":                           ["Philadelphia", "Pittsburgh", "Baltimore"],
  "Scranton":                             ["Philadelphia", "New York City", "Allentown"],
  "Baltimore":                            ["Philadelphia", "Northern Virginia", "Pittsburgh"],
  "Frederick":                            ["Baltimore", "Northern Virginia", "Philadelphia"],
  "Annapolis":                            ["Baltimore", "Northern Virginia", "Philadelphia"],
  "Wilmington":                           ["Philadelphia", "Baltimore", "Newark"],
  "Dover":                                ["Philadelphia", "Wilmington", "Baltimore"],
  "Newark":                               ["New York City", "Philadelphia", "Jersey City"],
  "Jersey City":                          ["Newark", "New York City", "Philadelphia"],
  "Trenton":                              ["Philadelphia", "Newark", "New York City"],
  "Camden":                               ["Philadelphia", "Wilmington", "Newark"],
  "New York City":                        ["Newark", "Jersey City", "Philadelphia", "Hartford", "Boston"],
  "Buffalo":                              ["Rochester", "Albany", "Pittsburgh", "Cleveland"],
  "Albany":                               ["Buffalo", "Hartford", "Boston", "New York City"],
  "Syracuse":                             ["Rochester", "Albany", "Buffalo"],
  "Hartford":                             ["New Haven", "Boston", "Providence", "New York City"],
  "New Haven":                            ["Hartford", "Bridgeport", "New York City"],
  "Bridgeport":                           ["New Haven", "Hartford", "New York City"],
  "Stamford":                             ["New York City", "Bridgeport", "Hartford"],
  "Boston":                               ["Worcester", "Providence", "Manchester", "Hartford"],
  "Worcester":                            ["Boston", "Providence", "Springfield"],
  "Providence":                           ["Boston", "Hartford", "Warwick"],
  "Warwick":                              ["Providence", "Boston", "Cranston"],
  "Cranston":                             ["Providence", "Warwick", "Boston"],
  "Manchester":                           ["Nashua", "Boston", "Providence"],
  "Nashua":                               ["Manchester", "Boston", "Worcester"],
  "Concord":                              ["Manchester", "Boston", "Burlington"],
  "Burlington":                           ["Montpelier", "Manchester", "Albany"],
  "Montpelier":                           ["Burlington", "Manchester", "Albany"],
  "Rutland":                              ["Burlington", "Albany", "Montpelier"],
  "Portland":                             ["Bangor", "Manchester", "Boston"],
  "Bangor":                               ["Portland", "Manchester", "Burlington"],
  "Augusta":                              ["Portland", "Bangor", "Manchester"],
  "Charleston":                           ["Huntington", "Morgantown", "Columbus", "Pittsburgh"],
  "Huntington":                           ["Charleston", "Louisville", "Cincinnati"],
  "Morgantown":                           ["Pittsburgh", "Charleston", "Columbus"],
  "Anchorage":                            ["Fairbanks", "Seattle"],
  "Fairbanks":                            ["Anchorage", "Seattle"],
  "Juneau":                               ["Anchorage", "Seattle"],
  "Honolulu":                             ["Maui", "Los Angeles"],
  "Maui":                                 ["Honolulu", "Hilo"],
  "Hilo":                                 ["Honolulu", "Maui"],
};

// ─── Fallback Resolution Logic ────────────────────────────────────────────────
// Given a selected metro and the full items array, returns:
// { resolvedMetro, isFallback, fallbackReason }
function resolveMetro(selectedMetro, items) {
  if (!selectedMetro) return { resolvedMetro: null, isFallback: false, fallbackReason: null };

  const hasData = (metro) => items.some(i => i.region === metro);

  if (hasData(selectedMetro)) {
    return { resolvedMetro: selectedMetro, isFallback: false, fallbackReason: null };
  }

  // Walk the fallback chain
  const chain = METRO_FALLBACK[selectedMetro] || [];
  for (const fallback of chain) {
    if (hasData(fallback)) {
      const isIntraState = Object.values(STATE_METROS).some(metros =>
        metros.includes(selectedMetro) && metros.includes(fallback)
      );
      return {
        resolvedMetro: fallback,
        isFallback: true,
        fallbackReason: isIntraState
          ? `No data for ${selectedMetro} — showing nearest available in state: ${fallback}`
          : `No data for ${selectedMetro} — showing nearest available market: ${fallback}`,
      };
    }
  }

  return {
    resolvedMetro: null,
    isFallback: true,
    fallbackReason: `No data available for ${selectedMetro} or any nearby markets.`,
  };
}

// ─── Metro Picker Component ───────────────────────────────────────────────────
function MetroPicker({ state, items, selectedMetro, onSelectMetro }) {
  if (!state) return null;
  const metros = STATE_METROS[state] || [];
  if (metros.length <= 1) return null; // no picker needed for single-metro states

  return (
    <div style={{
      backgroundColor: C.white,
      borderBottom: `1px solid ${C.border}`,
      padding: "8px 24px",
      display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap",
      flexShrink: 0,
    }}>
      <span style={{
        fontSize: 11, fontWeight: 600, color: C.textMuted,
        textTransform: "uppercase", letterSpacing: "0.08em",
        marginRight: 4, whiteSpace: "nowrap",
      }}>
        Metro
      </span>

      {/* All metros chip */}
      <button
        style={{
          height: 26, padding: "0 10px", borderRadius: 3,
          fontSize: 12, fontWeight: 500, cursor: "pointer",
          fontFamily: "inherit", whiteSpace: "nowrap",
          border: !selectedMetro ? `1px solid ${C.accentLight}` : `1px solid ${C.borderDark}`,
          backgroundColor: !selectedMetro ? "#EFF6FF" : C.white,
          color: !selectedMetro ? C.accentLight : C.textMuted,
        }}
        onClick={() => onSelectMetro(null)}
      >
        All metros
      </button>

      {metros.map(metro => {
        const count = items.filter(i => i.region === metro).length;
        const isSelected = selectedMetro === metro;
        return (
          <button
            key={metro}
            style={{
              height: 26, padding: "0 10px", borderRadius: 3,
              fontSize: 12, fontWeight: 500, cursor: "pointer",
              fontFamily: "inherit", whiteSpace: "nowrap",
              border: isSelected ? `1px solid ${C.accentLight}` : `1px solid ${C.borderDark}`,
              backgroundColor: isSelected ? "#EFF6FF" : C.white,
              color: isSelected ? C.accentLight : count > 0 ? C.textPrimary : C.textLight,
              display: "inline-flex", alignItems: "center", gap: 5,
            }}
            onClick={() => onSelectMetro(isSelected ? null : metro)}
          >
            {metro}
            {count > 0 && (
              <span style={{
                fontSize: 10, fontWeight: 700,
                backgroundColor: isSelected ? C.accentLight : C.neutralBg,
                color: isSelected ? "#fff" : C.textMuted,
                borderRadius: 2, padding: "0 4px", lineHeight: "16px",
              }}>
                {count}
              </span>
            )}
          </button>
        );
      })}
    </div>
  );
}


// ─── US Map Component (D3 Albers USA + real TopoJSON) ────────────────────────
// Uses d3-geo Albers USA projection + us-atlas TopoJSON from CDN.
// All 50 states rendered from real geographic boundaries.

function USMap({ items, selectedState, onSelectState }) {
  const [hoveredState, setHoveredState] = useState(null);
  const [mapOpen, setMapOpen]           = useState(true);
  const [statePaths, setStatePaths]     = useState({});   // { stateName: svgPath }
  const [stateCentroids, setStateCentroids] = useState({}); // { stateName: [cx, cy] }
  const [loading, setLoading]           = useState(true);
  const svgRef                          = useRef(null);

  const W = 960, H = 600;

  // FIPS id → state name lookup
  const FIPS_TO_STATE = {
    "01":"Alabama","02":"Alaska","04":"Arizona","05":"Arkansas","06":"California",
    "08":"Colorado","09":"Connecticut","10":"Delaware","12":"Florida","13":"Georgia",
    "15":"Hawaii","16":"Idaho","17":"Illinois","18":"Indiana","19":"Iowa",
    "20":"Kansas","21":"Kentucky","22":"Louisiana","23":"Maine","24":"Maryland",
    "25":"Massachusetts","26":"Michigan","27":"Minnesota","28":"Mississippi",
    "29":"Missouri","30":"Montana","31":"Nebraska","32":"Nevada","33":"New Hampshire",
    "34":"New Jersey","35":"New Mexico","36":"New York","37":"North Carolina",
    "38":"North Dakota","39":"Ohio","40":"Oklahoma","41":"Oregon","42":"Pennsylvania",
    "44":"Rhode Island","45":"South Carolina","46":"South Dakota","47":"Tennessee",
    "48":"Texas","49":"Utah","50":"Vermont","51":"Virginia","53":"Washington",
    "54":"West Virginia","55":"Wisconsin","56":"Wyoming",
  };

  // Item counts per state
  const itemCountByState = useMemo(() => {
    const counts = {};
    Object.entries(STATE_METROS).forEach(([state, metros]) => {
      counts[state] = items.filter(i => metros.includes(i.region)).length;
    });
    return counts;
  }, [items]);

  // Load TopoJSON + compute paths using d3
  useEffect(() => {
    let cancelled = false;
    async function load() {
      try {
        // Load d3-geo and topojson-client from CDN via dynamic import equivalent
        const [topoRes, d3Res] = await Promise.all([
          fetch("https://cdn.jsdelivr.net/npm/us-atlas@3/states-10m.json"),
          fetch("https://cdnjs.cloudflare.com/ajax/libs/d3/7.9.0/d3.min.js"),
        ]);
        if (cancelled) return;

        const topology = await topoRes.json();

        // Load d3 script into page if not already present
        if (!window.d3) {
          await new Promise((resolve, reject) => {
            const script = document.createElement("script");
            script.src = "https://cdnjs.cloudflare.com/ajax/libs/d3/7.9.0/d3.min.js";
            script.onload = resolve;
            script.onerror = reject;
            document.head.appendChild(script);
          });
        }
        if (!window.topojson) {
          await new Promise((resolve, reject) => {
            const script = document.createElement("script");
            script.src = "https://cdnjs.cloudflare.com/ajax/libs/topojson/3.0.2/topojson.min.js";
            script.onload = resolve;
            script.onerror = reject;
            document.head.appendChild(script);
          });
        }
        if (cancelled) return;

        const projection = window.d3.geoAlbersUsa().scale(1280).translate([W / 2, H / 2]);
        const path = window.d3.geoPath().projection(projection);
        const features = window.topojson.feature(topology, topology.objects.states).features;

        const paths = {};
        const centroids = {};
        features.forEach(feature => {
          const fips = String(feature.id).padStart(2, "0");
          const name = FIPS_TO_STATE[fips];
          if (!name) return;
          paths[name] = path(feature);
          const c = path.centroid(feature);
          if (c && !isNaN(c[0])) centroids[name] = c;
        });

        if (!cancelled) {
          setStatePaths(paths);
          setStateCentroids(centroids);
          setLoading(false);
        }
      } catch (err) {
        console.error("Map load error:", err);
        if (!cancelled) setLoading(false);
      }
    }
    load();
    return () => { cancelled = true; };
  }, []);

  function getStateFill(name) {
    if (selectedState === name) return "#2563EB";
    if (hoveredState === name) return "#BFDBFE";
    if (itemCountByState[name] > 0) return "#DBEAFE";
    return "#E2E8F0";
  }

  function getStateStroke(name) {
    return selectedState === name ? "#1D4ED8" : "#94A3B8";
  }

  function getStrokeWidth(name) {
    return selectedState === name ? 1.5 : 0.5;
  }

  function handleClick(name) {
    onSelectState(selectedState === name ? null : name);
  }

  const totalFiltered = selectedState
    ? items.filter(i => (STATE_METROS[selectedState] || []).includes(i.region)).length
    : items.length;

  const stateNames = Object.keys(statePaths);

  return (
    <div style={{ backgroundColor: C.white, borderBottom: `1px solid ${C.border}`, flexShrink: 0 }}>

      {/* Header / toggle */}
      <div
        style={{ display: "flex", alignItems: "center", padding: "8px 24px",
          borderBottom: mapOpen ? `1px solid ${C.border}` : "none",
          cursor: "pointer", userSelect: "none", backgroundColor: C.white }}
        onClick={() => setMapOpen(o => !o)}
      >
        <span style={{ fontSize: 11, fontWeight: 600, color: C.textMuted,
          textTransform: "uppercase", letterSpacing: "0.08em", flex: 1 }}>
          Geographic Filter
          {selectedState ? (
            <span style={{ marginLeft: 10, padding: "2px 8px", borderRadius: 3,
              backgroundColor: "#2563EB", color: "#fff",
              fontSize: 11, fontWeight: 600, textTransform: "none", letterSpacing: 0 }}>
              {selectedState} — {totalFiltered} item{totalFiltered !== 1 ? "s" : ""}
            </span>
          ) : (
            <span style={{ marginLeft: 10, color: C.textLight, fontWeight: 400,
              textTransform: "none", letterSpacing: 0, fontSize: 11 }}>
              All regions · {totalFiltered} items
            </span>
          )}
        </span>
        {selectedState && (
          <span style={{ fontSize: 12, color: "#2563EB", marginRight: 16, fontWeight: 500 }}
            onClick={e => { e.stopPropagation(); onSelectState(null); }}>
            Clear
          </span>
        )}
        <span style={{ fontSize: 11, color: C.textMuted }}>
          {mapOpen ? "▲ Collapse" : "▼ Expand"}
        </span>
      </div>

      {/* Map body */}
      {mapOpen && (
        <div style={{ position: "relative", padding: "8px 20px 4px" }}>
          {loading ? (
            <div style={{ height: 220, display: "flex", alignItems: "center",
              justifyContent: "center", color: C.textMuted, fontSize: 13 }}>
              Loading map…
            </div>
          ) : (
            <svg
              ref={svgRef}
              viewBox={`0 0 ${W} ${H}`}
              style={{ width: "100%", height: 220, display: "block" }}
              onMouseLeave={() => { setHoveredState(null); setTooltip(null); }}
            >
              {stateNames.map(name => (
                <g key={name}>
                  <path
                    d={statePaths[name]}
                    fill={getStateFill(name)}
                    stroke={getStateStroke(name)}
                    strokeWidth={getStrokeWidth(name)}
                    style={{ cursor: "pointer", transition: "fill 0.1s" }}
                    onClick={() => handleClick(name)}
                    onMouseEnter={() => setHoveredState(name)}
                    onMouseLeave={() => setHoveredState(null)}
                  />
                  {stateCentroids[name] && (
                    <text
                      x={stateCentroids[name][0]}
                      y={stateCentroids[name][1]}
                      fontSize="10"
                      fontFamily="Inter, system-ui, sans-serif"
                      fontWeight="600"
                      fill={selectedState === name ? "#fff" : "#475569"}
                      textAnchor="middle"
                      dominantBaseline="middle"
                      style={{ pointerEvents: "none", userSelect: "none" }}
                    >
                      {name === "Rhode Island" || name === "Delaware" ||
                       name === "Connecticut" || name === "New Jersey" ||
                       name === "Maryland" || name === "Massachusetts" ? null :
                        (() => {
                          const abbr = Object.entries({
                            "Alabama":"AL","Alaska":"AK","Arizona":"AZ","Arkansas":"AR",
                            "California":"CA","Colorado":"CO","Connecticut":"CT","Delaware":"DE",
                            "Florida":"FL","Georgia":"GA","Hawaii":"HI","Idaho":"ID",
                            "Illinois":"IL","Indiana":"IN","Iowa":"IA","Kansas":"KS",
                            "Kentucky":"KY","Louisiana":"LA","Maine":"ME","Maryland":"MD",
                            "Massachusetts":"MA","Michigan":"MI","Minnesota":"MN","Mississippi":"MS",
                            "Missouri":"MO","Montana":"MT","Nebraska":"NE","Nevada":"NV",
                            "New Hampshire":"NH","New Jersey":"NJ","New Mexico":"NM","New York":"NY",
                            "North Carolina":"NC","North Dakota":"ND","Ohio":"OH","Oklahoma":"OK",
                            "Oregon":"OR","Pennsylvania":"PA","Rhode Island":"RI","South Carolina":"SC",
                            "South Dakota":"SD","Tennessee":"TN","Texas":"TX","Utah":"UT",
                            "Vermont":"VT","Virginia":"VA","Washington":"WA","West Virginia":"WV",
                            "Wisconsin":"WI","Wyoming":"WY"
                          }).find(([k]) => k === name);
                          return abbr ? abbr[1] : null;
                        })()
                      }
                    </text>
                  )}
                </g>
              ))}
            </svg>
          )}

          {/* Legend */}
          <div style={{ display: "flex", alignItems: "center", gap: 16,
            padding: "2px 4px 6px", justifyContent: "flex-end" }}>
            {[
              { color: "#2563EB", label: "Selected" },
              { color: "#DBEAFE", label: "Has data" },
              { color: "#E2E8F0", label: "No data" },
            ].map(({ color, label }) => (
              <div key={label} style={{ display: "flex", alignItems: "center",
                gap: 5, fontSize: 11, color: C.textMuted }}>
                <div style={{ width: 12, height: 12, borderRadius: 2,
                  backgroundColor: color, border: `1px solid ${C.borderDark}` }} />
                {label}
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}



// ─── Bulk Import Modal ────────────────────────────────────────────────────────
function BulkImportModal({ existingItems, onClose, onImport }) {
  const [stage, setStage]         = useState("upload"); // upload | preview | done
  const [dragOver, setDragOver]   = useState(false);
  const [parseError, setParseError] = useState("");
  const [previewRows, setPreviewRows] = useState([]); // { raw, validated, status, conflict }
  const [importing, setImporting] = useState(false);
  const [importResult, setImportResult] = useState(null);
  const fileRef = useRef(null);

  function processFile(file) {
    setParseError("");
    const ext = file.name.split(".").pop().toLowerCase();
    if (!["csv","xlsx","xls"].includes(ext)) {
      setParseError("Only CSV and Excel files are supported.");
      return;
    }
    if (ext === "csv") {
      const reader = new FileReader();
      reader.onload = e => processText(e.target.result);
      reader.readAsText(file);
    } else {
      // XLSX via SheetJS
      const reader = new FileReader();
      reader.onload = e => {
        try {
          const wb  = XLSX.read(e.target.result, { type: "array" });
          const ws  = wb.Sheets[wb.SheetNames[0]];
          const csv = XLSX.utils.sheet_to_csv(ws);
          processText(csv);
        } catch(err) {
          setParseError("Could not parse Excel file: " + err.message);
        }
      };
      reader.readAsArrayBuffer(file);
    }
  }

  function processText(text) {
    const result = parseCSV(text);
    if (result.error) { setParseError(result.error); return; }
    const rows = result.rows.map((raw, idx) => {
      const { errors, warnings, conflict } = validateRow(raw, existingItems);
      let status = "new";
      if (errors.length)   status = "error";
      else if (conflict)   status = "conflict";
      else if (warnings.length) status = "warning";
      return { idx, raw, errors, warnings, conflict, status };
    });
    setPreviewRows(rows);
    setStage("preview");
  }

  function handleDrop(e) {
    e.preventDefault(); setDragOver(false);
    const file = e.dataTransfer?.files?.[0];
    if (file) processFile(file);
  }

  function handleFileSelect(e) {
    const file = e.target.files?.[0];
    if (file) processFile(file);
    e.target.value = "";
  }

  function handleConfirmImport() {
    setImporting(true);
    const importable = previewRows.filter(r => r.status !== "error");
    const newItems = [], updated = [], skipped = [];
    importable.forEach(({ raw, conflict }) => {
      const item = {
        id: conflict ? conflict.id : generateId(),
        org_id: "org-1",
        csi_code:    raw.csi_code.trim(),
        description: raw.description.trim(),
        unit:        raw.unit.trim().toUpperCase(),
        mat_cost:    Math.round((parseFloat(raw.mat_cost)  || 0) * 100),
        lab_cost:    Math.round((parseFloat(raw.lab_cost)  || 0) * 100),
        equip_cost:  Math.round((parseFloat(raw.equip_cost)|| 0) * 100),
        region:      raw.region?.trim() || "—",
        source:      raw.source?.trim() || "Other",
        notes:       raw.notes?.trim() || "",
        updated_by:  MOCK_USER.initials,
        updated_at:  new Date().toISOString().slice(0, 10),
      };
      if (conflict) updated.push(item);
      else          newItems.push(item);
    });
    const errored = previewRows.filter(r => r.status === "error");
    setTimeout(() => {
      onImport(newItems, updated);
      setImportResult({ newItems: newItems.length, updated: updated.length, skipped: errored.length });
      setImporting(false);
      setStage("done");
    }, 400);
  }

  const counts = {
    total:    previewRows.length,
    new:      previewRows.filter(r => r.status === "new").length,
    conflict: previewRows.filter(r => r.status === "conflict").length,
    warning:  previewRows.filter(r => r.status === "warning").length,
    error:    previewRows.filter(r => r.status === "error").length,
  };

  const statusStyle = (status) => ({
    new:      { bg: "#DCFCE7", color: "#16A34A", label: "NEW" },
    conflict: { bg: "#FEF9C3", color: "#854D0E", label: "OVERWRITE" },
    warning:  { bg: "#FEF3C7", color: "#B45309", label: "WARNING" },
    error:    { bg: "#FEE2E2", color: "#DC2626", label: "ERROR" },
  }[status] || { bg: C.neutralBg, color: C.textMuted, label: status.toUpperCase() });

  return (
    <div style={s.overlay} onClick={e => e.target === e.currentTarget && onClose()}>
      <div style={{ ...s.modal(true), width: stage === "preview" ? "88vw" : 560, maxWidth: 1100 }}>

        {/* Header */}
        <div style={s.modalHeader}>
          <div>
            <div style={s.modalTitle}>
              {stage === "upload" && "Bulk Import Cost Items"}
              {stage === "preview" && `Preview Import — ${counts.total} row${counts.total !== 1 ? "s" : ""}`}
              {stage === "done" && "Import Complete"}
            </div>
            <div style={s.modalSubtitle}>
              {stage === "upload" && "CSV or Excel · Must match template column format"}
              {stage === "preview" && `${counts.new} new · ${counts.conflict} overwrites · ${counts.error} errors`}
              {stage === "done" && `${importResult?.newItems} added · ${importResult?.updated} updated · ${importResult?.skipped} skipped`}
            </div>
          </div>
          <button style={s.closeBtn} onClick={onClose}>×</button>
        </div>

        {/* ── UPLOAD STAGE ── */}
        {stage === "upload" && (
          <div style={s.modalBody}>
            {/* Template download */}
            <div style={{
              display: "flex", alignItems: "center", justifyContent: "space-between",
              padding: "10px 14px", backgroundColor: "#EFF6FF", borderRadius: 4,
              border: "1px solid #BFDBFE",
            }}>
              <div>
                <div style={{ fontSize: 13, fontWeight: 600, color: C.textPrimary }}>Download Template</div>
                <div style={{ fontSize: 11, color: C.textMuted, marginTop: 2 }}>
                  CSV with required columns and sample rows. Fill it in and upload below.
                </div>
              </div>
              <button style={{ ...s.btn("primary"), height: 30, fontSize: 12 }} onClick={downloadTemplate}>
                ↓ Template CSV
              </button>
            </div>

            {/* Required columns list */}
            <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
              {IMPORT_TEMPLATE_COLS.map(col => (
                <span key={col} style={{
                  padding: "2px 8px", borderRadius: 3, fontSize: 11, fontWeight: 600,
                  backgroundColor: C.neutralBg, color: C.textMuted,
                  border: `1px solid ${C.borderDark}`, fontFamily: "monospace",
                }}>{col}</span>
              ))}
            </div>

            {/* Drop zone */}
            {parseError && <div style={s.errorMsg}>{parseError}</div>}
            <div
              style={{
                border: `2px dashed ${dragOver ? C.accentLight : C.borderDark}`,
                borderRadius: 4, padding: "32px 24px",
                display: "flex", flexDirection: "column", alignItems: "center", gap: 10,
                backgroundColor: dragOver ? "#EFF6FF" : C.contentBg,
                transition: "all 0.1s", cursor: "pointer", textAlign: "center",
              }}
              onClick={() => fileRef.current?.click()}
              onDragOver={e => { e.preventDefault(); setDragOver(true); }}
              onDragLeave={() => setDragOver(false)}
              onDrop={handleDrop}
            >
              <div style={{ fontSize: 32 }}>📂</div>
              <div style={{ fontSize: 14, fontWeight: 600, color: C.textPrimary }}>
                Drag & drop your file here
              </div>
              <div style={{ fontSize: 12, color: C.textMuted }}>or click to browse</div>
              <div style={{ display: "flex", gap: 6, marginTop: 4 }}>
                {["CSV", "XLSX", "XLS"].map(fmt => (
                  <span key={fmt} style={{
                    padding: "2px 8px", borderRadius: 3, fontSize: 11, fontWeight: 600,
                    border: `1px solid ${C.borderDark}`, color: C.textMuted, backgroundColor: C.white,
                  }}>{fmt}</span>
                ))}
              </div>
              <input ref={fileRef} type="file" accept=".csv,.xlsx,.xls"
                style={{ display: "none" }} onChange={handleFileSelect} />
            </div>
          </div>
        )}

        {/* ── PREVIEW STAGE ── */}
        {stage === "preview" && (
          <>
            {/* Status summary strip */}
            <div style={{
              display: "flex", gap: 0,
              borderBottom: `1px solid ${C.border}`,
              flexShrink: 0,
            }}>
              {[
                { label: "New",       count: counts.new,      bg: "#DCFCE7", color: "#16A34A" },
                { label: "Overwrite", count: counts.conflict,  bg: "#FEF9C3", color: "#854D0E" },
                { label: "Warning",   count: counts.warning,   bg: "#FEF3C7", color: "#B45309" },
                { label: "Error",     count: counts.error,     bg: "#FEE2E2", color: "#DC2626" },
              ].map(({ label, count, bg, color }) => (
                <div key={label} style={{
                  padding: "10px 20px", borderRight: `1px solid ${C.border}`,
                  backgroundColor: count > 0 ? bg : C.white,
                  flex: 1, textAlign: "center",
                }}>
                  <div style={{ fontSize: 20, fontWeight: 700, color: count > 0 ? color : C.textLight }}>
                    {count}
                  </div>
                  <div style={{ fontSize: 11, color: count > 0 ? color : C.textLight, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.06em" }}>
                    {label}
                  </div>
                </div>
              ))}
            </div>

            {/* Preview table */}
            <div style={{ overflowY: "auto", flex: 1 }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                <thead style={{ position: "sticky", top: 0, backgroundColor: C.white, zIndex: 1 }}>
                  <tr>
                    {["Status", "CSI Code", "Description", "Unit", "Material", "Labor", "Equipment", "Total", "Region", "Source", "Notes / Issues"].map(h => (
                      <th key={h} style={{
                        padding: "7px 10px", textAlign: "left",
                        borderBottom: `2px solid ${C.border}`,
                        fontSize: 10, fontWeight: 600, color: C.textMuted,
                        textTransform: "uppercase", letterSpacing: "0.06em",
                        whiteSpace: "nowrap",
                      }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {previewRows.map(({ idx, raw, errors, warnings, conflict, status }) => {
                    const st = statusStyle(status);
                    const mat   = parseFloat(raw.mat_cost)  || 0;
                    const lab   = parseFloat(raw.lab_cost)  || 0;
                    const equip = parseFloat(raw.equip_cost)|| 0;
                    const total = mat + lab + equip;
                    const issues = [...errors, ...warnings,
                      ...(conflict ? [`Overwrites: ${conflict.description}`] : [])
                    ];
                    return (
                      <tr key={idx} style={{ backgroundColor: idx % 2 === 1 ? C.tableStripe : C.white }}>
                        <td style={{ padding: "6px 10px", whiteSpace: "nowrap" }}>
                          <span style={{
                            padding: "2px 7px", borderRadius: 3, fontSize: 10, fontWeight: 700,
                            backgroundColor: st.bg, color: st.color,
                            textTransform: "uppercase", letterSpacing: "0.05em",
                          }}>{st.label}</span>
                        </td>
                        <td style={{ padding: "6px 10px", fontFamily: "monospace", color: C.textMuted, whiteSpace: "nowrap" }}>{raw.csi_code}</td>
                        <td style={{ padding: "6px 10px", color: C.textPrimary, maxWidth: 240, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{raw.description}</td>
                        <td style={{ padding: "6px 10px", fontFamily: "monospace", color: C.textMuted }}>{raw.unit}</td>
                        <td style={{ padding: "6px 10px", textAlign: "right", fontFamily: "monospace", color: C.textMuted }}>${mat.toFixed(2)}</td>
                        <td style={{ padding: "6px 10px", textAlign: "right", fontFamily: "monospace", color: C.textMuted }}>${lab.toFixed(2)}</td>
                        <td style={{ padding: "6px 10px", textAlign: "right", fontFamily: "monospace", color: C.textMuted }}>${equip.toFixed(2)}</td>
                        <td style={{ padding: "6px 10px", textAlign: "right", fontFamily: "monospace", fontWeight: 700, color: C.textPrimary }}>${total.toFixed(2)}</td>
                        <td style={{ padding: "6px 10px", color: C.textMuted, whiteSpace: "nowrap" }}>{raw.region}</td>
                        <td style={{ padding: "6px 10px", color: C.textMuted, whiteSpace: "nowrap" }}>{raw.source}</td>
                        <td style={{ padding: "6px 10px", fontSize: 11, color: issues.length ? "#DC2626" : C.textLight, maxWidth: 220, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
                          {issues.length ? issues.join(" · ") : raw.notes || "—"}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>

            {counts.error > 0 && (
              <div style={{
                padding: "8px 16px", backgroundColor: "#FEF2F2",
                borderTop: `1px solid #FECACA`, fontSize: 12, color: "#991B1B", flexShrink: 0,
              }}>
                ⚠ {counts.error} row{counts.error !== 1 ? "s" : ""} with errors will be skipped.
                {counts.new + counts.conflict + counts.warning > 0 &&
                  ` ${counts.new + counts.conflict + counts.warning} rows will be imported.`}
              </div>
            )}
          </>
        )}

        {/* ── DONE STAGE ── */}
        {stage === "done" && (
          <div style={{ ...s.modalBody, alignItems: "center", textAlign: "center", padding: "40px 32px" }}>
            <div style={{ fontSize: 40, marginBottom: 12 }}>✅</div>
            <div style={{ fontSize: 16, fontWeight: 700, color: C.textPrimary, marginBottom: 8 }}>
              Import Complete
            </div>
            <div style={{ fontSize: 13, color: C.textMuted, lineHeight: 1.8 }}>
              <div><strong style={{ color: "#16A34A" }}>{importResult?.newItems}</strong> new cost items added</div>
              <div><strong style={{ color: "#854D0E" }}>{importResult?.updated}</strong> existing items updated</div>
              {importResult?.skipped > 0 && <div><strong style={{ color: "#DC2626" }}>{importResult?.skipped}</strong> rows skipped (errors)</div>}
            </div>
          </div>
        )}

        {/* Footer */}
        <div style={{ ...s.modalFooter, justifyContent: "space-between" }}>
          <div>
            {stage === "preview" && (
              <button style={s.btn()} onClick={() => { setStage("upload"); setPreviewRows([]); setParseError(""); }}>
                ← Back
              </button>
            )}
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            {stage !== "done" && <button style={s.btn()} onClick={onClose}>Cancel</button>}
            {stage === "preview" && (
              <button
                style={{ ...s.btn("primary"), opacity: importing ? 0.7 : 1 }}
                onClick={handleConfirmImport}
                disabled={importing || counts.new + counts.conflict + counts.warning === 0}
              >
                {importing ? "Importing…" : `Import ${counts.new + counts.conflict + counts.warning} Row${counts.new + counts.conflict + counts.warning !== 1 ? "s" : ""}`}
              </button>
            )}
            {stage === "done" && <button style={s.btn("primary")} onClick={onClose}>Done</button>}
          </div>
        </div>
      </div>
    </div>
  );
}

// ─── Division Accordion ───────────────────────────────────────────────────────
function DivisionAccordion({ items, selectedState, onSelectSub, onImport, allItems }) {
  const [openDivision, setOpenDivision] = useState(null);
  const [hoverDiv, setHoverDiv]         = useState(null);
  const [hoverSub, setHoverSub]         = useState(null);
  const [search, setSearch]             = useState("");
  const [showImport, setShowImport]     = useState(false);

  // items already pre-filtered by CostDatabase (state + metro + fallback)
  const filteredItems = items;

  const filtered = CSI_STRUCTURE.filter(div =>
    !search ||
    div.label.toLowerCase().includes(search.toLowerCase()) ||
    `division ${div.code}`.includes(search.toLowerCase()) ||
    div.subdivisions.some(s => s.label.toLowerCase().includes(search.toLowerCase()))
  );

  function toggleDiv(code) {
    setOpenDivision(prev => prev === code ? null : code);
  }

  return (
    <>
      {/* Toolbar */}
      <div style={s.toolbar}>
        <input
          style={s.searchInput}
          placeholder="Filter divisions or subdivisions…"
          value={search}
          onChange={e => setSearch(e.target.value)}
        />
        <div style={s.spacer} />
        <span style={{ fontSize: 12, color: C.textMuted, marginRight: 12 }}>
          {filteredItems.length} cost item{filteredItems.length !== 1 ? "s" : ""} · {CSI_STRUCTURE.length} divisions
          {selectedState && <span style={{ color: C.accentLight, marginLeft: 4 }}>· {selectedState}</span>}
        </span>
        <button style={s.btn()} onClick={downloadTemplate}>↓ Template</button>
        <button style={s.btn("primary")} onClick={() => setShowImport(true)}>↑ Import</button>
      </div>

      {showImport && (
        <BulkImportModal
          existingItems={allItems}
          onClose={() => setShowImport(false)}
          onImport={(newItems, updated) => {
            onImport(newItems, updated);
            setShowImport(false);
          }}
        />
      )}

      {/* Accordion list */}
      <div style={s.accordionWrap}>
        {filtered.map(div => {
          const isOpen   = openDivision === div.code;
          const divCount = filteredItems.filter(i => i.csi_code.startsWith(div.code)).length;

          return (
            <div key={div.code}>
              {/* Division row */}
              <div
                style={s.divRow(isOpen, hoverDiv === div.code)}
                onMouseEnter={() => setHoverDiv(div.code)}
                onMouseLeave={() => setHoverDiv(null)}
                onClick={() => toggleDiv(div.code)}
              >
                <span style={s.divRowArrow(isOpen)}>▶</span>
                <span style={s.divRowCode}>{div.code}</span>
                <span style={s.divRowLabel(isOpen)}>{div.code} — {div.label}</span>
                <span style={s.divRowCount}>
                  {divCount > 0 ? `${divCount} item${divCount !== 1 ? "s" : ""}` : "—"}
                </span>
              </div>

              {/* Subdivision rows */}
              {isOpen && (
                <div style={s.subPanel}>
                  {div.subdivisions.map(sub => {
                    const subCount = filteredItems.filter(i => i.csi_code === sub.code).length;
                    const hKey = `${div.code}-${sub.code}`;
                    return (
                      <div
                        key={sub.code}
                        style={s.subAccRow(hoverSub === hKey)}
                        onMouseEnter={() => setHoverSub(hKey)}
                        onMouseLeave={() => setHoverSub(null)}
                        onClick={() => onSelectSub(sub, div)}
                      >
                        <span style={s.subAccArrow}>▶</span>
                        <span style={s.subAccCode}>{sub.code}</span>
                        <span style={s.subAccLabel}>{sub.label}</span>
                        <span style={s.subAccCount}>
                          {subCount > 0
                            ? `${subCount} item${subCount !== 1 ? "s" : ""}`
                            : <span style={{ color: C.textLight }}>no items</span>}
                        </span>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          );
        })}
      </div>
    </>
  );
}

// ─── Quote Upload Panel ───────────────────────────────────────────────────────
// QuoteUploadPanel — compact upload strip only. Quote table shown per-subdivision in LineItemsTable.
function QuoteUploadPanel({ quotes, onAddQuote, costItems, onCreateCostItem }) {
  const [dragOver, setDragOver] = useState(false);
  const [showModal, setShowModal] = useState(false);
  const fileRef = useRef(null);

  function readFileForModal(file) {
    const fileType = file.name.split(".").pop().toUpperCase();
    if (fileType === "PDF") {
      // Read as base64 data URL — works inside sandboxed iframes
      const reader = new FileReader();
      reader.onload = e => {
        setShowModal({ fileName: file.name, fileType, fileUrl: e.target.result });
      };
      reader.readAsDataURL(file);
    } else {
      setShowModal({ fileName: file.name, fileType, fileUrl: null });
    }
  }

  function handleFileDrop(e) {
    e.preventDefault();
    setDragOver(false);
    const file = e.dataTransfer?.files?.[0];
    if (file) readFileForModal(file);
  }

  function handleFileSelect(e) {
    const file = e.target.files?.[0];
    if (file) readFileForModal(file);
    e.target.value = "";
  }

  return (
    <>
      <div style={{
        backgroundColor: C.white,
        borderBottom: `1px solid ${C.border}`,
        padding: "10px 24px",
        flexShrink: 0,
      }}>
        <div
          style={{
            border: `2px dashed ${dragOver ? C.accentLight : C.borderDark}`,
            borderRadius: 4,
            padding: "10px 16px",
            display: "flex", alignItems: "center", gap: 14,
            backgroundColor: dragOver ? "#EFF6FF" : C.contentBg,
            transition: "all 0.1s",
            cursor: "pointer",
          }}
          onClick={() => fileRef.current?.click()}
          onDragOver={e => { e.preventDefault(); setDragOver(true); }}
          onDragLeave={() => setDragOver(false)}
          onDrop={handleFileDrop}
        >
          <div style={{
            width: 32, height: 32, borderRadius: 3,
            backgroundColor: C.neutralBg,
            display: "flex", alignItems: "center", justifyContent: "center",
            fontSize: 16, flexShrink: 0,
          }}>📄</div>
          <div style={{ flex: 1 }}>
            <div style={{ fontSize: 12, fontWeight: 600, color: C.textPrimary }}>
              Upload Subcontractor Quote
            </div>
            <div style={{ fontSize: 11, color: C.textMuted, marginTop: 1 }}>
              Drag & drop or click to browse — then map line items to CSI codes after upload.
              {quotes.length > 0 && (
                <span style={{ marginLeft: 8, color: C.accentLight, fontWeight: 600 }}>
                  {quotes.length} quote{quotes.length !== 1 ? "s" : ""} on file — visible when you open a subdivision.
                </span>
              )}
            </div>
          </div>
          <div style={{ display: "flex", gap: 5, flexShrink: 0 }}>
            {["XLSX", "CSV", "PDF"].map(fmt => (
              <span key={fmt} style={{
                padding: "2px 7px", borderRadius: 3, fontSize: 11, fontWeight: 600,
                border: `1px solid ${C.borderDark}`, color: C.textMuted,
                backgroundColor: C.white,
              }}>{fmt}</span>
            ))}
          </div>
          <input
            ref={fileRef} type="file"
            accept=".xlsx,.csv,.pdf"
            style={{ display: "none" }}
            onChange={handleFileSelect}
          />
        </div>
      </div>

      {showModal && (
        <QuoteMapModal
          fileName={showModal.fileName}
          fileType={showModal.fileType}
          fileUrl={showModal.fileUrl}
          costItems={costItems}
          onClose={() => setShowModal(false)}
          onSave={(quotes) => { quotes.forEach(q => onAddQuote(q)); setShowModal(false); }}
          onCreateCostItem={onCreateCostItem}
        />
      )}
    </>
  );
}

// ─── PDF Viewer Modal ────────────────────────────────────────────────────────
function PDFViewerModal({ quote, onClose }) {
  // For real uploaded files: render via object URL in an iframe.
  // For mock/pre-existing quotes without a file_url: render a realistic mock proposal.
  const hasPdf = !!quote.file_url;

  const fmtMoney = (cents) =>
    "$" + (cents / 100).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

  const mockHtml = `
    <html>
    <head>
      <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: Arial, sans-serif; font-size: 13px; color: #1e293b; padding: 48px 56px; background: #fff; }
        .header { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 36px; padding-bottom: 20px; border-bottom: 2px solid #e2e8f0; }
        .company-name { font-size: 22px; font-weight: 700; color: #0f172a; }
        .company-sub { font-size: 12px; color: #64748b; margin-top: 4px; }
        .doc-title { font-size: 28px; font-weight: 700; color: #0f172a; margin-bottom: 24px; }
        .meta-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 28px; }
        .meta-block label { font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.08em; color: #94a3b8; display: block; margin-bottom: 3px; }
        .meta-block span { font-size: 13px; color: #0f172a; }
        .section-title { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.08em; color: #64748b; border-bottom: 1px solid #e2e8f0; padding-bottom: 6px; margin-bottom: 12px; margin-top: 28px; }
        table { width: 100%; border-collapse: collapse; }
        th { text-align: left; font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.07em; color: #64748b; padding: 8px 10px; border-bottom: 2px solid #e2e8f0; }
        th.right { text-align: right; }
        td { padding: 9px 10px; border-bottom: 1px solid #f1f5f9; font-size: 13px; color: #0f172a; }
        td.right { text-align: right; font-family: monospace; }
        td.muted { color: #64748b; }
        .total-row td { font-weight: 700; border-top: 2px solid #e2e8f0; border-bottom: none; font-size: 14px; }
        .notes-box { background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 4px; padding: 14px 16px; margin-top: 12px; font-size: 12px; color: #475569; line-height: 1.6; }
        .footer { margin-top: 48px; padding-top: 20px; border-top: 1px solid #e2e8f0; display: flex; justify-content: space-between; font-size: 11px; color: #94a3b8; }
        .status-badge { display: inline-block; padding: 3px 10px; border-radius: 3px; font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.05em;
          background: ${ quote.status === "Approved" ? "#dcfce7" : quote.status === "Reviewed" ? "#f5f3ff" : "#eff6ff" };
          color: ${ quote.status === "Approved" ? "#16a34a" : quote.status === "Reviewed" ? "#6d28d9" : "#1d4ed8" }; }
        .validity { background: #fef9c3; border: 1px solid #fde047; border-radius: 3px; padding: 8px 14px; font-size: 12px; color: #854d0e; margin-top: 20px; }
      </style>
    </head>
    <body>
      <div class="header">
        <div>
          <div class="company-name">${quote.company}</div>
          <div class="company-sub">${quote.trade} Contractor &nbsp;|&nbsp; ${quote.market}</div>
          ${quote.contact_name ? '<div class="company-sub" style="margin-top:6px;">' + quote.contact_name + (quote.email ? " &nbsp;·&nbsp; " + quote.email : "") + (quote.phone ? " &nbsp;·&nbsp; " + quote.phone : "") + "</div>" : ""}
        </div>
        <div style="text-align:right;">
          <span class="status-badge">${quote.status}</span>
          <div style="font-size:11px;color:#94a3b8;margin-top:8px;">Quote #${quote.id}</div>
        </div>
      </div>

      <div class="doc-title">Subcontractor Proposal</div>

      <div class="meta-grid">
        <div class="meta-block"><label>CSI Division</label><span>${quote.csi_code} — ${quote.csi_label}</span></div>
        <div class="meta-block"><label>Market / Region</label><span>${quote.market}</span></div>
        <div class="meta-block"><label>Quote Date</label><span>${quote.quote_date}</span></div>
        <div class="meta-block"><label>Valid Through</label><span>${quote.valid_through || "Not specified"}</span></div>
      </div>

      <div class="section-title">Scope of Work</div>
      <table>
        <thead>
          <tr>
            <th>Description</th>
            <th>Unit</th>
            <th class="right">Unit Cost</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td>${quote.description || "Per scope documents"}</td>
            <td class="muted">${quote.unit}</td>
            <td class="right">${fmtMoney(quote.unit_cost)}</td>
          </tr>
        </tbody>
        <tfoot>
          <tr class="total-row">
            <td colspan="2">Unit Cost Total</td>
            <td class="right">${fmtMoney(quote.unit_cost)}</td>
          </tr>
        </tfoot>
      </table>

      ${quote.notes ? '<div class="section-title">Clarifications &amp; Exclusions</div><div class="notes-box">' + quote.notes + '</div>' : ""}

      <div class="validity">
        ⚠ This proposal is valid through <strong>${quote.valid_through || "date not specified"}</strong>.
        Pricing subject to change after expiration. Re-quote required for scope changes.
      </div>

      <div class="footer">
        <span>File: ${quote.file_name}</span>
        <span>Uploaded ${quote.uploaded_at} by ${quote.uploaded_by} &nbsp;·&nbsp; Meridian Construction Group</span>
      </div>
    </body>
    </html>
  `;

  return (
    <div style={{
      position: "fixed", inset: 0,
      backgroundColor: "rgba(0,0,0,0.6)",
      display: "flex", alignItems: "center", justifyContent: "center",
      zIndex: 200,
    }}
      onClick={e => e.target === e.currentTarget && onClose()}
    >
      <div style={{
        backgroundColor: C.white,
        borderRadius: 4,
        width: "88vw", height: "90vh",
        display: "flex", flexDirection: "column",
        boxShadow: "0 16px 48px rgba(0,0,0,0.3)",
        overflow: "hidden",
      }}>
        {/* Viewer header */}
        <div style={{
          display: "flex", alignItems: "center", padding: "12px 20px",
          borderBottom: `1px solid ${C.border}`,
          backgroundColor: C.white, flexShrink: 0,
        }}>
          <div style={{ flex: 1 }}>
            <div style={{ fontSize: 14, fontWeight: 600, color: C.textPrimary }}>
              {quote.file_name}
            </div>
            <div style={{ fontSize: 11, color: C.textMuted, marginTop: 2 }}>
              {quote.company} · {quote.csi_code} — {quote.csi_label}
            </div>
          </div>
          <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
            {!hasPdf && (
              <span style={{
                fontSize: 11, color: C.textMuted, fontStyle: "italic", marginRight: 8,
              }}>
                Mock preview — upload a real file to view the original
              </span>
            )}
            <button
              style={{
                ...{height: 30, padding: "0 14px", fontSize: 12, borderRadius: 3,
                cursor: "pointer", fontFamily: "inherit", fontWeight: 500,
                border: `1px solid ${C.borderDark}`, backgroundColor: C.white,
                color: C.textPrimary},
              }}
              onClick={onClose}
            >
              Close
            </button>
          </div>
        </div>

        {/* Document viewer */}
        <div style={{ flex: 1, overflow: "hidden", backgroundColor: "#64748B" }}>
          {hasPdf ? (
            <iframe
              src={quote.file_url}
              style={{ width: "100%", height: "100%", border: "none" }}
              title={quote.file_name}
            />
          ) : (
            <iframe
              srcDoc={mockHtml}
              style={{ width: "100%", height: "100%", border: "none", backgroundColor: "#fff" }}
              title={quote.file_name}
            />
          )}
        </div>
      </div>
    </div>
  );
}

// ─── Subdivision Quote Panel — shown inside line-items view ───────────────────
function SubdivisionQuotePanel({ quotes, csiCode, onUpdateQuotes }) {
  const [open, setOpen] = useState(true);
  const [viewingQuote, setViewingQuote] = useState(null);
  const matchingQuotes = quotes.filter(q => q.csi_code === csiCode);

  const statusColors = {
    Received: { bg: "#EFF6FF", color: "#1D4ED8" },
    Reviewed: { bg: "#F5F3FF", color: "#6D28D9" },
    Approved: { bg: "#DCFCE7", color: "#16A34A" },
    Rejected: { bg: "#FEE2E2", color: "#DC2626" },
  };

  if (matchingQuotes.length === 0) return null;

  return (
    <>
    <div style={{
      backgroundColor: C.white,
      borderBottom: `1px solid ${C.border}`,
      flexShrink: 0,
    }}>
      {/* Header */}
      <div
        style={{
          display: "flex", alignItems: "center", padding: "8px 24px",
          borderBottom: open ? `1px solid ${C.border}` : "none",
          cursor: "pointer", userSelect: "none",
        }}
        onClick={() => setOpen(o => !o)}
      >
        <span style={{
          fontSize: 11, fontWeight: 600, color: C.textMuted,
          textTransform: "uppercase", letterSpacing: "0.08em", flex: 1,
        }}>
          Subcontractor Quotes
          <span style={{
            marginLeft: 10, padding: "2px 7px", borderRadius: 3,
            backgroundColor: "#EFF6FF", color: C.accentLight,
            fontSize: 11, fontWeight: 600, textTransform: "none", letterSpacing: 0,
          }}>
            {matchingQuotes.length} quote{matchingQuotes.length !== 1 ? "s" : ""} for this subdivision
          </span>
        </span>
        <span style={{ fontSize: 11, color: C.textMuted }}>
          {open ? "▲ Collapse" : "▼ Expand"}
        </span>
      </div>

      {open && (
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
            <thead>
              <tr>
                {["Company", "Description", "Unit", "Unit Cost", "Market", "Quote Date", "Valid Through", "Status", "File", ""].map(h => (
                  <th key={h} style={{
                    padding: "6px 14px", textAlign: h === "Unit Cost" ? "right" : "left",
                    borderBottom: `2px solid ${C.border}`,
                    fontSize: 10, fontWeight: 600, color: C.textMuted,
                    textTransform: "uppercase", letterSpacing: "0.07em",
                    whiteSpace: "nowrap", backgroundColor: C.white,
                  }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {matchingQuotes.map((q, i) => (
                <tr key={q.id} style={{ backgroundColor: i % 2 === 1 ? C.tableStripe : C.white }}>
                  <td style={{ padding: "7px 14px", whiteSpace: "nowrap" }}>
                    <div style={{ fontWeight: 500, color: C.textPrimary }}>{q.company}</div>
                    <div style={{ fontSize: 11, color: C.textMuted }}>{q.contact_name}</div>
                  </td>
                  <td style={{ padding: "7px 14px", color: C.textPrimary, maxWidth: 260, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{q.description}</td>
                  <td style={{ padding: "7px 14px", color: C.textMuted, fontFamily: "monospace" }}>{q.unit}</td>
                  <td style={{ padding: "7px 14px", textAlign: "right", fontFamily: "monospace", fontWeight: 700, color: C.textPrimary, whiteSpace: "nowrap" }}>
                    ${(q.unit_cost / 100).toLocaleString("en-US", { minimumFractionDigits: 2 })}
                  </td>
                  <td style={{ padding: "7px 14px", color: C.textMuted, whiteSpace: "nowrap" }}>{q.market}</td>
                  <td style={{ padding: "7px 14px", color: C.textMuted, whiteSpace: "nowrap" }}>{q.quote_date}</td>
                  <td style={{ padding: "7px 14px", color: C.textMuted, whiteSpace: "nowrap" }}>{q.valid_through || "—"}</td>
                  <td style={{ padding: "7px 14px", whiteSpace: "nowrap" }}>
                    <span style={{
                      padding: "2px 7px", borderRadius: 3, fontSize: 10, fontWeight: 700,
                      textTransform: "uppercase", letterSpacing: "0.05em",
                      backgroundColor: (statusColors[q.status] || statusColors.Received).bg,
                      color: (statusColors[q.status] || statusColors.Received).color,
                    }}>{q.status}</span>
                  </td>
                  <td style={{ padding: "7px 14px", whiteSpace: "nowrap" }}>
                    <span style={{
                      padding: "2px 6px", borderRadius: 2, fontSize: 10, fontWeight: 600,
                      border: `1px solid ${C.borderDark}`, color: C.textMuted,
                    }}>{q.file_type}</span>
                    <span
                      style={{
                        marginLeft: 6, fontSize: 11,
                        color: C.accentLight, cursor: "pointer",
                        textDecoration: "underline",
                      }}
                      onClick={e => { e.stopPropagation(); setViewingQuote(q); }}
                    >
                      {q.file_name.length > 24 ? q.file_name.slice(0, 24) + "…" : q.file_name}
                    </span>
                  </td>
                  <td style={{ padding: "7px 14px", textAlign: "right" }}>
                    <select
                      value={q.status}
                      onChange={e => {
                        const updated = quotes.map(x => x.id === q.id ? { ...x, status: e.target.value } : x);
                        onUpdateQuotes(updated);
                      }}
                      onClick={e => e.stopPropagation()}
                      style={{
                        fontSize: 11, fontWeight: 600, padding: "2px 6px",
                        border: `1px solid ${C.borderDark}`, borderRadius: 3,
                        color: { Received: C.accentLight, Reviewed: "#6D28D9", Approved: "#16A34A", Rejected: "#DC2626" }[q.status] || C.textMuted,
                        backgroundColor: C.white, fontFamily: "inherit", cursor: "pointer",
                      }}
                    >
                      {QUOTE_STATUSES.map(st => <option key={st} value={st}>{st}</option>)}
                    </select>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>

      {viewingQuote && (
        <PDFViewerModal
          quote={viewingQuote}
          onClose={() => setViewingQuote(null)}
        />
      )}
    </>
  );
}

// ─── Quote Status Dropdown ────────────────────────────────────────────────────
function QuoteStatusDropdown({ quote, onChange }) {
  const colors = {
    Received: C.accentLight,
    Reviewed: "#6D28D9",
    Approved: "#16A34A",
    Rejected: "#DC2626",
  };
  return (
    <select
      value={quote.status}
      onChange={e => onChange(quote.id, e.target.value)}
      onClick={e => e.stopPropagation()}
      style={{
        fontSize: 11, fontWeight: 600, padding: "2px 6px",
        border: `1px solid ${C.borderDark}`, borderRadius: 3,
        color: colors[quote.status] || C.textMuted,
        backgroundColor: C.white, fontFamily: "inherit", cursor: "pointer",
      }}
    >
      {QUOTE_STATUSES.map(s => <option key={s} value={s}>{s}</option>)}
    </select>
  );
}

// ─── Quote Map Modal ──────────────────────────────────────────────────────────
// ─── PDF.js Canvas Viewer ─────────────────────────────────────────────────────
// Renders PDF pages onto canvas elements using PDF.js — no iframe, no sandbox issues.
// PDFPageCanvas — renders a single PDF page onto a canvas
function PDFPageCanvas({ pdf, pageNum, containerWidth }) {
  const canvasRef = useRef(null);

  useEffect(() => {
    if (!pdf || !canvasRef.current) return;
    let cancelled = false;
    let renderTask = null;
    async function render() {
      try {
        const page = await pdf.getPage(pageNum);
        if (cancelled) return;
        const viewport = page.getViewport({ scale: 1 });
        const scale = Math.min((containerWidth - 32) / viewport.width, 1.6);
        const scaled = page.getViewport({ scale });
        const canvas = canvasRef.current;
        canvas.width  = scaled.width;
        canvas.height = scaled.height;
        renderTask = page.render({ canvasContext: canvas.getContext("2d"), viewport: scaled });
        await renderTask.promise;
      } catch (err) {
        if (err?.name !== "RenderingCancelledException") console.error(err);
      }
    }
    render();
    return () => { cancelled = true; if (renderTask) renderTask.cancel(); };
  }, [pdf, pageNum, containerWidth]);

  return (
    <canvas
      ref={canvasRef}
      style={{
        display: "block", maxWidth: "100%",
        boxShadow: "0 4px 16px rgba(0,0,0,0.4)", borderRadius: 2,
      }}
    />
  );
}

function PDFJSViewer({ dataUrl }) {
  const containerRef  = useRef(null);
  const [loading, setLoading] = useState(true);
  const [error, setError]     = useState(null);
  const [pdf, setPdf]         = useState(null);
  const [numPages, setNumPages] = useState(0);
  const [containerWidth, setContainerWidth] = useState(700);

  // Load PDF.js and parse the document
  useEffect(() => {
    let cancelled = false;
    async function load() {
      if (!window.pdfjsLib) {
        await new Promise((resolve, reject) => {
          const script = document.createElement("script");
          script.src = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js";
          script.onload = resolve;
          script.onerror = reject;
          document.head.appendChild(script);
        });
        window.pdfjsLib.GlobalWorkerOptions.workerSrc =
          "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
      }
      if (cancelled) return;
      try {
        const base64  = dataUrl.split(",")[1];
        const binary  = atob(base64);
        const bytes   = new Uint8Array(binary.length);
        for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
        const doc = await window.pdfjsLib.getDocument({ data: bytes }).promise;
        if (cancelled) return;
        setPdf(doc);
        setNumPages(doc.numPages);
        setLoading(false);
      } catch (err) {
        if (!cancelled) setError("Could not render PDF: " + err.message);
        setLoading(false);
      }
    }
    load();
    return () => { cancelled = true; };
  }, [dataUrl]);

  // Track container width for responsive scaling
  useEffect(() => {
    if (!containerRef.current) return;
    const ro = new ResizeObserver(entries => {
      for (const entry of entries) setContainerWidth(entry.contentRect.width);
    });
    ro.observe(containerRef.current);
    setContainerWidth(containerRef.current.clientWidth);
    return () => ro.disconnect();
  }, []);

  if (error) return (
    <div style={{ flex: 1, display: "flex", alignItems: "center", justifyContent: "center",
      color: "#FCA5A5", fontSize: 12, padding: 24, textAlign: "center" }}>
      ⚠ {error}
    </div>
  );

  return (
    <div ref={containerRef} style={{
      flex: 1, overflowY: "auto",
      display: "flex", flexDirection: "column",
      alignItems: "center", gap: 12,
      padding: "16px 16px 24px",
    }}>
      {loading && (
        <div style={{ color: "#94A3B8", fontSize: 13, marginTop: 40 }}>Loading PDF…</div>
      )}
      {!loading && Array.from({ length: numPages }, (_, i) => i + 1).map(pageNum => (
        <div key={pageNum} style={{ width: "100%", display: "flex", flexDirection: "column", alignItems: "center", gap: 4 }}>
          <PDFPageCanvas pdf={pdf} pageNum={pageNum} containerWidth={containerWidth} />
          {numPages > 1 && (
            <div style={{ fontSize: 10, color: "#64748B", letterSpacing: "0.05em" }}>
              PAGE {pageNum} OF {numPages}
            </div>
          )}
        </div>
      ))}
    </div>
  );
}

function QuoteMapModal({ fileName, fileType, fileUrl, costItems: initialCostItems, onClose, onSave, defaultCsiCode, onCreateCostItem }) {
  const today = new Date().toISOString().slice(0, 10);

  // Local copy of costItems so newly created items are immediately available in dropdowns
  const [localCostItems, setLocalCostItems] = React.useState(initialCostItems || []);

  // Shared header fields
  const [company,      setCompany]      = useState("");
  const [market,       setMarket]       = useState("");
  const [quoteDate,    setQuoteDate]    = useState(today);
  const [notes,        setNotes]        = useState("");
  const [error,        setError]        = useState("");

  // Line items array
  const blankRow = () => ({ _key: Math.random(), csi_code: defaultCsiCode || "", cost_item_id: "", description: "", unit: "SF", unit_cost: "", _newItemOpen: false, _newItem: { description: "", unit: "SF", mat_cost: "", lab_cost: "", equip_cost: "" } });
  const [lineItems, setLineItems] = useState([blankRow()]);

  const allSubs = CSI_STRUCTURE.flatMap(div =>
    div.subdivisions.map(sub => ({ code: sub.code, label: `${sub.code} — ${sub.label}` }))
  );

  function updateRow(idx, field, value) {
    setLineItems(prev => prev.map((row, i) => {
      if (i !== idx) return row;
      const updated = { ...row, [field]: value };
      if (field === "csi_code") { updated.cost_item_id = ""; updated._newItemOpen = false; }
      return updated;
    }));
  }

  function updateNewItem(idx, field, value) {
    setLineItems(prev => prev.map((row, i) =>
      i !== idx ? row : { ...row, _newItem: { ...row._newItem, [field]: value } }
    ));
  }

  function toggleNewItemForm(idx, open) {
    setLineItems(prev => prev.map((row, i) =>
      i !== idx ? row : { ...row, _newItemOpen: open, _newItem: { description: "", unit: "SF", mat_cost: "", lab_cost: "", equip_cost: "" } }
    ));
  }

  function commitNewItem(idx) {
    const row = lineItems[idx];
    const ni  = row._newItem;
    if (!ni.description.trim()) return;

    const newItem = {
      id: generateId(),
      org_id: "org-1",
      csi_code: row.csi_code,
      description: ni.description.trim(),
      unit: ni.unit,
      mat_cost:   Math.round((parseFloat(ni.mat_cost)  || 0) * 100),
      lab_cost:   Math.round((parseFloat(ni.lab_cost)  || 0) * 100),
      equip_cost: Math.round((parseFloat(ni.equip_cost)|| 0) * 100),
      region: market || "",
      source: "Bid",
      notes: "",
      updated_by: MOCK_USER.name.split(" ").map(n => n[0]).join(". ") + ".",
      updated_at: today,
    };

    // Persist to the cost database
    if (onCreateCostItem) onCreateCostItem(newItem);

    // Add to local items so the dropdown reflects it immediately
    setLocalCostItems(prev => [...prev, newItem]);

    // Select the new item in this row and close the mini-form
    setLineItems(prev => prev.map((r, i) =>
      i !== idx ? r : { ...r, cost_item_id: newItem.id, _newItemOpen: false, description: newItem.description, unit: newItem.unit }
    ));
  }

  function addRow() { setLineItems(prev => [...prev, blankRow()]); }
  function removeRow(idx) { setLineItems(prev => prev.length === 1 ? prev : prev.filter((_, i) => i !== idx)); }

  function handleSave() {
    if (!company.trim()) { setError("Company name is required."); return; }
    for (let i = 0; i < lineItems.length; i++) {
      const row = lineItems[i];
      if (!row.csi_code)     { setError(`Row ${i + 1}: CSI subdivision is required.`); return; }
      if (!row.cost_item_id) { setError(`Row ${i + 1}: Line item is required.`); return; }
      if (!row.unit_cost)    { setError(`Row ${i + 1}: Unit cost is required.`); return; }
    }
    const sharedMeta = {
      org_id: "org-1",
      company: company.trim(),
      market,
      quote_date: quoteDate,
      file_name: fileName,
      file_type: fileType,
      file_url: fileUrl || null,
      notes: notes.trim(),
      status: "Received",
      uploaded_by: MOCK_USER.name.split(" ").map(n => n[0]).join(". ") + ".",
      uploaded_at: today,
    };
    const quotes = lineItems.map(row => {
      const subLabel = allSubs.find(s => s.code === row.csi_code)?.label || "";
      return {
        ...sharedMeta,
        id: generateQuoteId(),
        cost_item_id: row.cost_item_id,
        csi_code: row.csi_code,
        csi_label: subLabel.split(" — ")[1] || "",
        description: row.description.trim(),
        unit: row.unit,
        unit_cost: Math.round((parseFloat(row.unit_cost) || 0) * 100),
        total_amount: 0,
        trade: "",
        contact_name: "",
        email: "",
        phone: "",
      };
    });
    onSave(quotes);
  }

  return (
    <div style={s.overlay} onClick={e => e.target === e.currentTarget && onClose()}>
      <div style={{
        backgroundColor: C.white, borderRadius: 4,
        width: "92vw", maxWidth: 1300, height: "90vh",
        display: "flex", flexDirection: "column",
        boxShadow: "0 16px 48px rgba(0,0,0,0.25)",
        overflow: "hidden",
      }}>

        {/* Header */}
        <div style={{ ...s.modalHeader, flexShrink: 0, borderBottom: `1px solid ${C.border}` }}>
          <div>
            <div style={s.modalTitle}>Map Quote to Line Items</div>
            <div style={s.modalSubtitle}>
              <span style={{
                padding: "1px 6px", borderRadius: 2, fontSize: 11, fontWeight: 600,
                border: `1px solid ${C.borderDark}`, color: C.textMuted, marginRight: 8,
              }}>{fileType}</span>
              {fileName}
            </div>
          </div>
          <button style={s.closeBtn} onClick={onClose}>×</button>
        </div>

        {/* Two-panel body */}
        <div style={{ display: "flex", flex: 1, overflow: "hidden" }}>

          {/* LEFT — Form */}
          <div style={{ flex: "0 0 45%", display: "flex", flexDirection: "column", overflow: "hidden", borderRight: `1px solid ${C.border}` }}>
            <div style={{ ...s.modalBody, flex: 1, overflowY: "auto" }}>
              {error && <div style={s.errorMsg}>{error}</div>}

              <div style={s.sectionDivider}>Subcontractor</div>
              <div style={s.field}>
                <label style={s.label}>Company Name *</label>
                <input style={s.input} value={company} onChange={e => setCompany(e.target.value)} placeholder="e.g. Lone Star Concrete LLC" autoFocus />
              </div>

              <div style={s.sectionDivider}>Quote Info</div>
              <StateMetroSelect value={market} onChange={setMarket} />
              <div style={s.grid3}>
                <div style={s.field}>
                  <label style={s.label}>Quote Date</label>
                  <input style={s.input} type="date" value={quoteDate} onChange={e => setQuoteDate(e.target.value)} />
                </div>

              </div>
              <div style={s.field}>
                <label style={s.label}>Notes</label>
                <textarea style={s.textarea} value={notes} onChange={e => setNotes(e.target.value)} placeholder="Scope inclusions, exclusions, lead times, warranty notes…" />
              </div>

              {/* Line Items */}
              <div style={{ ...s.sectionDivider, display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                <span>Line Items</span>
                <button style={{ ...s.btn("primary"), height: 24, fontSize: 11, padding: "0 10px" }} onClick={addRow}>
                  + Add Line Item
                </button>
              </div>

              <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
                {lineItems.map((row, idx) => {
                  const rowItems = localCostItems.filter(i => i.csi_code === row.csi_code);
                  const ni = row._newItem;
                  const niTotal = (parseFloat(ni.mat_cost)||0) + (parseFloat(ni.lab_cost)||0) + (parseFloat(ni.equip_cost)||0);

                  return (
                    <div key={row._key} style={{
                      border: `1px solid ${row._newItemOpen ? C.accentLight : C.border}`,
                      borderRadius: 4, padding: "12px 12px 10px",
                      backgroundColor: C.contentBg,
                    }}>
                      {/* Row header */}
                      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10 }}>
                        <span style={{ fontSize: 11, fontWeight: 700, color: C.textMuted, textTransform: "uppercase", letterSpacing: "0.07em" }}>
                          Line Item {idx + 1}
                        </span>
                        {lineItems.length > 1 && (
                          <button onClick={() => removeRow(idx)} style={{ background: "none", border: "none", cursor: "pointer", color: C.danger, fontSize: 16, lineHeight: 1, padding: 0 }}>×</button>
                        )}
                      </div>

                      {/* CSI + Line Item picker */}
                      <div style={{ ...s.grid2, marginBottom: 8 }}>
                        <div style={s.field}>
                          <label style={s.label}>CSI Subdivision *</label>
                          <select style={s.fieldSelect} value={row.csi_code} onChange={e => updateRow(idx, "csi_code", e.target.value)}>
                            <option value="">— Select subdivision —</option>
                            {CSI_STRUCTURE.map(div => (
                              <optgroup key={div.code} label={`Division ${div.code} — ${div.label}`}>
                                {div.subdivisions.map(sub => (
                                  <option key={sub.code} value={sub.code}>{sub.code} — {sub.label}</option>
                                ))}
                              </optgroup>
                            ))}
                          </select>
                        </div>
                        <div style={s.field}>
                          <label style={{ ...s.label, display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                            <span>Line Item *</span>
                            {row.csi_code && !row._newItemOpen && (
                              <span
                                onClick={() => toggleNewItemForm(idx, true)}
                                style={{ fontSize: 11, color: C.accentLight, cursor: "pointer", fontWeight: 600, textDecoration: "none" }}
                              >
                                + New Cost Item
                              </span>
                            )}
                          </label>
                          {!row._newItemOpen ? (
                            <>
                              <select
                                style={s.fieldSelect}
                                value={row.cost_item_id}
                                onChange={e => updateRow(idx, "cost_item_id", e.target.value)}
                                disabled={!row.csi_code}
                              >
                                <option value="">{row.csi_code ? "— Select line item —" : "— Select subdivision first —"}</option>
                                {rowItems.map(i => (
                                  <option key={i.id} value={i.id}>{i.description}</option>
                                ))}
                              </select>
                              {row.csi_code && rowItems.length === 0 && (
                                <div style={{ fontSize: 11, color: C.textMuted, marginTop: 3 }}>
                                  No items yet —{" "}
                                  <span style={{ color: C.accentLight, cursor: "pointer" }} onClick={() => toggleNewItemForm(idx, true)}>
                                    create one
                                  </span>
                                </div>
                              )}
                            </>
                          ) : (
                            <div style={{ fontSize: 11, color: C.accentLight, fontStyle: "italic", height: 34, display: "flex", alignItems: "center" }}>
                              Creating new cost item below…
                            </div>
                          )}
                        </div>
                      </div>

                      {/* Inline new cost item form */}
                      {row._newItemOpen && (
                        <div style={{
                          marginTop: 4, marginBottom: 8,
                          padding: "12px",
                          backgroundColor: "#EFF6FF",
                          border: `1px solid #BFDBFE`,
                          borderRadius: 3,
                        }}>
                          <div style={{ fontSize: 11, fontWeight: 700, color: C.accentLight, textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 10 }}>
                            New Cost Item — {row.csi_code}
                          </div>
                          <div style={{ display: "grid", gridTemplateColumns: "1fr 80px", gap: 8, marginBottom: 8 }}>
                            <div style={s.field}>
                              <label style={s.label}>Description *</label>
                              <input
                                style={s.input}
                                value={ni.description}
                                onChange={e => updateNewItem(idx, "description", e.target.value)}
                                placeholder="e.g. 5″ PCC Slab on Grade"
                                autoFocus
                              />
                            </div>
                            <div style={s.field}>
                              <label style={s.label}>Unit</label>
                              <select style={s.fieldSelect} value={ni.unit} onChange={e => updateNewItem(idx, "unit", e.target.value)}>
                                {UNITS.map(u => <option key={u} value={u}>{u}</option>)}
                              </select>
                            </div>
                          </div>
                          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8, marginBottom: 10 }}>
                            {[["mat_cost","Material"],["lab_cost","Labor"],["equip_cost","Equipment"]].map(([key, lbl]) => (
                              <div key={key} style={s.field}>
                                <label style={s.label}>{lbl} ($/unit)</label>
                                <input
                                  style={{ ...s.input, fontFamily: "monospace" }}
                                  type="number" min="0" step="0.01"
                                  value={ni[key]}
                                  onChange={e => updateNewItem(idx, key, e.target.value)}
                                  placeholder="0.00"
                                />
                              </div>
                            ))}
                          </div>
                          {niTotal > 0 && (
                            <div style={{ fontSize: 11, color: C.textMuted, marginBottom: 10, fontFamily: "monospace" }}>
                              Total unit cost: <strong style={{ color: C.textPrimary }}>${niTotal.toFixed(2)}</strong>
                            </div>
                          )}
                          <div style={{ display: "flex", gap: 8 }}>
                            <button
                              style={{ ...s.btn("primary"), height: 28, fontSize: 12 }}
                              onClick={() => commitNewItem(idx)}
                              disabled={!ni.description.trim()}
                            >
                              Add to Cost Book &amp; Select
                            </button>
                            <button style={{ ...s.btn(), height: 28, fontSize: 12 }} onClick={() => toggleNewItemForm(idx, false)}>
                              Cancel
                            </button>
                          </div>
                        </div>
                      )}

                      {/* Description + Unit + Unit Cost */}
                      <div style={{ display: "grid", gridTemplateColumns: "1fr 80px 110px", gap: 8 }}>
                        <div style={s.field}>
                          <label style={s.label}>Description</label>
                          <input style={s.input} value={row.description} onChange={e => updateRow(idx, "description", e.target.value)} placeholder="e.g. 5″ PCC Slab" />
                        </div>
                        <div style={s.field}>
                          <label style={s.label}>Unit</label>
                          <select style={s.fieldSelect} value={row.unit} onChange={e => updateRow(idx, "unit", e.target.value)}>
                            {UNITS.map(u => <option key={u} value={u}>{u}</option>)}
                          </select>
                        </div>
                        <div style={s.field}>
                          <label style={s.label}>Unit Cost *</label>
                          <input
                            style={{ ...s.input, fontFamily: "monospace" }}
                            type="number" min="0" step="0.01"
                            value={row.unit_cost}
                            onChange={e => updateRow(idx, "unit_cost", e.target.value)}
                            placeholder="0.00"
                          />
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>

            {/* Footer */}
            <div style={{ ...s.modalFooter, justifyContent: "space-between", borderTop: `1px solid ${C.border}`, flexShrink: 0 }}>
              <span style={{ fontSize: 12, color: C.textMuted }}>
                {lineItems.length} line item{lineItems.length !== 1 ? "s" : ""}
              </span>
              <div style={{ display: "flex", gap: 8 }}>
                <button style={s.btn()} onClick={onClose}>Cancel</button>
                <button style={s.btn("primary")} onClick={handleSave}>
                  Save {lineItems.length > 1 ? `${lineItems.length} Quotes` : "Quote"}
                </button>
              </div>
            </div>
          </div>

          {/* RIGHT — PDF Viewer */}
          <div style={{ flex: "0 0 55%", display: "flex", flexDirection: "column", overflow: "hidden", backgroundColor: "#475569" }}>
            <div style={{ padding: "6px 16px", backgroundColor: "#334155", fontSize: 11, color: "#94A3B8", fontWeight: 500, flexShrink: 0, display: "flex", alignItems: "center", gap: 8 }}>
              <span>📄</span>
              <span style={{ color: "#CBD5E1" }}>{fileName}</span>
            </div>
            {fileUrl && fileType === "PDF"
              ? <PDFJSViewer dataUrl={fileUrl} />
              : (
                <div style={{ flex: 1, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 12, color: "#64748B" }}>
                  <div style={{ fontSize: 48 }}>📄</div>
                  <div style={{ fontSize: 14, fontWeight: 600, color: "#CBD5E1" }}>{fileName}</div>
                  <div style={{ fontSize: 12, textAlign: "center", lineHeight: 1.6 }}>
                    PDF preview available for .pdf uploads.<br/>
                    Fill in the form on the left to map this quote.
                  </div>
                </div>
              )
            }
          </div>
        </div>
      </div>
    </div>
  );
}

// ─── Cost Database Root ───────────────────────────────────────────────────────
function CostDatabase() {
  const [items, setItems]                 = useState(MOCK_ITEMS);
  const [view, setView]                   = useState("accordion");
  const [activeDivision, setDiv]          = useState(null);
  const [activeSub, setSub]               = useState(null);
  const [editModal, setEditModal]         = useState(null);
  const [selectedState, setSelectedState] = useState(null);
  const [selectedMetro, setSelectedMetro] = useState(null);
  const [quotes, setQuotes]               = useState(MOCK_QUOTES);

  function handleAddQuote(quoteOrUpdated, mode) {
    if (mode === "update") { setQuotes(quoteOrUpdated); return; }
    setQuotes(prev => [quoteOrUpdated, ...prev]);
  }

  // Selecting a new state resets metro selection
  function handleSelectState(state) {
    setSelectedState(state);
    setSelectedMetro(null);
  }

  function handleSelectSub(sub, div) { setSub(sub); setDiv(div); setView("line-items"); }
  function handleBackToAccordion()   { setSub(null); setDiv(null); setView("accordion"); }
  function handleAddItem(sub)        { setEditModal({ item: null, subdivision: sub }); }
  function handleEditItem(item, sub) { setEditModal({ item, subdivision: sub }); }

  function handleSaveItem(saved, isEdit) {
    setItems(prev => isEdit
      ? prev.map(i => i.id === saved.id ? saved : i)
      : [...prev, saved]
    );
  }

  function handleDeleteItem(id) {
    setItems(prev => prev.filter(i => i.id !== id));
  }

  function handleImport(newItems, updated) {
    setItems(prev => {
      // Apply updates to existing items
      const updatedMap = new Map(updated.map(u => [u.id, u]));
      const merged = prev.map(i => updatedMap.has(i.id) ? updatedMap.get(i.id) : i);
      // Add new items
      return [...merged, ...newItems];
    });
  }

  // Resolve metro fallback
  const { resolvedMetro, isFallback, fallbackReason } = useMemo(
    () => resolveMetro(selectedMetro, items),
    [selectedMetro, items]
  );

  // Items visible given current state + metro filters (with fallback applied)
  const visibleItems = useMemo(() => {
    if (!selectedState) return items;
    const stateMetros = STATE_METROS[selectedState] || [];
    if (!selectedMetro) return items.filter(i => stateMetros.includes(i.region));
    if (resolvedMetro)  return items.filter(i => i.region === resolvedMetro);
    return []; // selected metro + no fallback found
  }, [items, selectedState, selectedMetro, resolvedMetro]);

  // Pass filter context down for accordion counts
  const accordionItems = useMemo(() => {
    if (!selectedState) return items;
    const stateMetros = STATE_METROS[selectedState] || [];
    if (!selectedMetro) return items.filter(i => stateMetros.includes(i.region));
    if (resolvedMetro)  return items.filter(i => i.region === resolvedMetro);
    return [];
  }, [items, selectedState, selectedMetro, resolvedMetro]);

  return (
    <>
      {/* US Map */}
      <USMap
        items={items}
        selectedState={selectedState}
        onSelectState={handleSelectState}
      />

      {/* Metro picker — shown when a state with multiple metros is selected */}
      <MetroPicker
        state={selectedState}
        items={items}
        selectedMetro={selectedMetro}
        onSelectMetro={setSelectedMetro}
      />

      {/* Quote Upload Panel */}
      <QuoteUploadPanel
        quotes={quotes}
        onAddQuote={handleAddQuote}
        costItems={items}
        onCreateCostItem={item => { handleSaveItem(item, false); return item; }}
      />

      {view === "accordion" && (
        <DivisionAccordion
          items={accordionItems}
          allItems={items}
          selectedState={selectedState}
          onSelectSub={handleSelectSub}
          onImport={handleImport}
        />
      )}

      {view === "line-items" && activeSub && (
        <>
          <div style={s.breadcrumb}>
            <span style={s.crumbLink} onClick={handleBackToAccordion}>Cost Database</span>
            <span style={s.crumbSep}>›</span>
            <span style={{ ...s.crumbCurrent, color: C.textMuted }}>
              Division {activeDivision?.code} — {activeDivision?.label}
            </span>
            <span style={s.crumbSep}>›</span>
            <span style={s.crumbCurrent}>{activeSub.code} — {activeSub.label}</span>
            {selectedState && (
              <>
                <span style={s.crumbSep}>·</span>
                <span style={{ fontSize: 12, color: C.accentLight, fontWeight: 500 }}>
                  {selectedState}{resolvedMetro ? ` › ${resolvedMetro}` : ""}
                </span>
              </>
            )}
          </div>

          {/* Fallback banner */}
          {isFallback && fallbackReason && (
            <div style={{
              display: "flex", alignItems: "center", gap: 10,
              padding: "8px 24px",
              backgroundColor: "#FEF9C3",
              borderBottom: "1px solid #FDE047",
              fontSize: 12, color: "#854D0E", flexShrink: 0,
            }}>
              <span style={{ fontSize: 14 }}>⚠</span>
              <span>{fallbackReason}</span>
            </div>
          )}
          {isFallback && !fallbackReason && (
            <div style={{
              display: "flex", alignItems: "center", gap: 10,
              padding: "8px 24px",
              backgroundColor: "#FEE2E2",
              borderBottom: "1px solid #FECACA",
              fontSize: 12, color: "#991B1B", flexShrink: 0,
            }}>
              <span style={{ fontSize: 14 }}>⚠</span>
              <span>No data available for the selected market or any nearby alternatives.</span>
            </div>
          )}

          <LineItemsTable
            items={visibleItems}
            subdivision={activeSub}
            onAdd={handleAddItem}
            onEdit={handleEditItem}
            onDelete={handleDeleteItem}
            quotes={quotes}
            onUpdateQuotes={setQuotes}
          />
        </>
      )}

      {editModal && (
        <CostItemModal
          item={editModal.item}
          subdivisionCode={editModal.subdivision.code}
          subdivisionLabel={editModal.subdivision.label}
          onClose={() => setEditModal(null)}
          onSave={handleSaveItem}
          onDelete={handleDeleteItem}
        />
      )}
    </>
  );
}


// ─── Estimate Builder ─────────────────────────────────────────────────────────

let NEXT_PROJECT_ID = 1;
const generateProjectId = () => `PRJ-${String(++NEXT_PROJECT_ID).padStart(4,"0")}`;
let NEXT_EST_LINE_ID = 1;
const generateLineId = () => `EL-${++NEXT_EST_LINE_ID}`;

// ── Firm-level defaults (would live in org settings in production) ─────────────
const FIRM_DEFAULTS = {
  insSsRate:      0.28,   // 28% burden on direct labor
  feeRate:        0.05,   // 5% GC fee
  bondRate:       0.0182, // 1.82% bond
  glRate:         0.0098, // 0.98% GL insurance
  grtRate:        0.00,   // Gross receipts tax
  contingencyRate:0.00,
};

const SUPERVISION_ROLES = [
  { key: "project_executive",        label: "Project Executive",         defaultRate: 122.53 },
  { key: "senior_pm",                label: "Senior Project Manager",    defaultRate: 114.05 },
  { key: "project_manager",          label: "Project Manager",           defaultRate: 99.24  },
  { key: "asst_pm",                  label: "Assistant Project Manager", defaultRate: 68.49  },
  { key: "senior_superintendent",    label: "Senior Superintendent",     defaultRate: 90.00  },
  { key: "superintendent_1",         label: "Superintendent 1",          defaultRate: 82.00  },
  { key: "superintendent_2",         label: "Superintendent 2",          defaultRate: 82.00  },
  { key: "asst_superintendent_1",    label: "Asst. Superintendent 1",    defaultRate: 75.00  },
  { key: "asst_superintendent_2",    label: "Asst. Superintendent 2",    defaultRate: 75.00  },
  { key: "project_engineer_nonbim",  label: "Project Engineer (Non-BIM)",defaultRate: 60.00  },
  { key: "project_engineer_bim",     label: "Project Engineer (BIM)",    defaultRate: 65.00  },
  { key: "project_accountant",       label: "Project Accountant",        defaultRate: 60.00  },
  { key: "field_engineer",           label: "Field Engineer",            defaultRate: 57.79  },
  { key: "rodman",                   label: "Rodman",                    defaultRate: 30.00  },
  { key: "bim_vdc_manager",          label: "BIM/VDC Manager",           defaultRate: 68.49  },
  { key: "bim_vdc_specialist",       label: "BIM/VDC Specialist",        defaultRate: 57.79  },
  { key: "scheduler",               label: "Scheduler",                 defaultRate: 88.32  },
  { key: "senior_estimator",        label: "Senior Estimator",          defaultRate: 99.24  },
  { key: "estimator",               label: "Estimator",                 defaultRate: 68.49  },
  { key: "safety_director",         label: "Safety Director",           defaultRate: 114.05 },
];

const BIDSHEET_TRADES = [
  // ── Division 01 — General Requirements ────────────────────────────────────
  { code: "01 00 00", item: "SUBGUARD TOTAL",                         type: "NA" },
  { code: "01 00 00", item: "SUBCONTRACTOR BONDS",                    type: "No" },
  { code: "01 45 00", item: "QUALITY CONTROL / TESTING & INSPECTION", type: "No" },
  // ── Division 02 — Existing Conditions ─────────────────────────────────────
  { code: "02 41 00", item: "DEMOLITION",                             type: "S"  },
  { code: "02 41 19", item: "SELECTIVE DEMOLITION",                   type: "S"  },
  // ── Division 03 — Concrete ────────────────────────────────────────────────
  { code: "03 20 00", item: "CONCRETE REINFORCING MATERIAL",          type: "M"  },
  { code: "03 30 00", item: "CAST-IN-PLACE CONCRETE (TURNKEY)",       type: "S"  },
  { code: "03 30 00", item: "CONCRETE MATERIAL — 3,000 PSI",          type: "M"  },
  { code: "03 30 00", item: "CONCRETE MATERIAL — 4,000 PSI",          type: "M"  },
  { code: "03 41 00", item: "PRECAST STRUCTURAL CONCRETE",            type: "S"  },
  { code: "03 54 00", item: "POLISHED CONCRETE",                       type: "S"  },
  // ── Division 04 — Masonry ─────────────────────────────────────────────────
  { code: "04 20 00", item: "UNIT MASONRY",                           type: "S"  },
  { code: "04 40 00", item: "STONE ASSEMBLIES",                       type: "S"  },
  // ── Division 05 — Metals ──────────────────────────────────────────────────
  { code: "05 12 00", item: "STRUCTURAL STEEL MATERIAL",               type: "M"  },
  { code: "05 12 00", item: "STRUCTURAL STEEL ERECTION",               type: "S"  },
  { code: "05 21 00", item: "STEEL JOIST FRAMING",                    type: "S"  },
  { code: "05 31 00", item: "STEEL DECKING",                          type: "S"  },
  { code: "05 50 00", item: "METAL FABRICATIONS / MISC METALS",       type: "M"  },
  { code: "05 51 00", item: "METAL STAIRS",                           type: "S"  },
  { code: "05 52 00", item: "PIPE & TUBE RAILINGS / ORNAMENTAL",      type: "S"  },
  // ── Division 06 — Wood, Plastics & Composites ─────────────────────────────
  { code: "06 10 00", item: "ROUGH CARPENTRY",                        type: "S"  },
  { code: "06 40 00", item: "ARCHITECTURAL WOODWORK / MILLWORK",      type: "S"  },
  // ── Division 07 — Thermal & Moisture Protection ───────────────────────────
  { code: "07 10 00", item: "DAMPPROOFING & WATERPROOFING",           type: "S"  },
  { code: "07 24 00", item: "EXTERIOR INSULATION & FINISH (EIFS)",    type: "S"  },
  { code: "07 25 00", item: "WEATHER BARRIERS",                       type: "S"  },
  { code: "07 40 00", item: "ROOFING & SIDING PANELS / METAL SIDING", type: "S"  },
  { code: "07 51 00", item: "BUILT-UP BITUMINOUS ROOFING",            type: "S"  },
  { code: "07 54 00", item: "THERMOPLASTIC MEMBRANE ROOFING",         type: "S"  },
  { code: "07 60 00", item: "FLASHING & SHEET METAL",                 type: "S"  },
  { code: "07 81 00", item: "APPLIED FIREPROOFING (SPRAYED)",         type: "S"  },
  { code: "07 84 00", item: "FIRESTOPPING",                           type: "S"  },
  { code: "07 92 00", item: "JOINT SEALANTS",                         type: "S"  },
  // ── Division 08 — Openings ────────────────────────────────────────────────
  { code: "08 11 00", item: "METAL DOORS & FRAMES",                   type: "S"  },
  { code: "08 14 00", item: "WOOD DOORS",                             type: "S"  },
  { code: "08 31 00", item: "ACCESS DOORS & PANELS",                  type: "S"  },
  { code: "08 33 00", item: "COILING DOORS & GRILLES / OVERHEAD",     type: "S"  },
  { code: "08 41 00", item: "ENTRANCES & STOREFRONTS",                type: "S"  },
  { code: "08 44 00", item: "CURTAIN WALL & GLAZED ASSEMBLIES",       type: "S"  },
  { code: "08 51 00", item: "METAL WINDOWS",                          type: "S"  },
  { code: "08 71 00", item: "DOOR HARDWARE",                          type: "S"  },
  { code: "08 80 00", item: "GLAZING / GLASS & MIRRORS",              type: "S"  },
  // ── Division 09 — Finishes ────────────────────────────────────────────────
  { code: "09 21 00", item: "GYPSUM BOARD ASSEMBLIES / DRYWALL",      type: "S"  },
  { code: "09 30 00", item: "TILING / CERAMIC & QUARRY TILE",         type: "S"  },
  { code: "09 51 00", item: "ACOUSTICAL CEILINGS",                    type: "S"  },
  { code: "09 65 00", item: "RESILIENT FLOORING",                     type: "S"  },
  { code: "09 68 00", item: "CARPET",                                 type: "S"  },
  { code: "09 90 00", item: "PAINTS & COATINGS / PAINTING",           type: "S"  },
  // ── Division 10 — Specialties ─────────────────────────────────────────────
  { code: "10 00 00", item: "SPECIALTIES (MISC)",                     type: "S"  },
  { code: "10 14 00", item: "SIGNAGE",                                type: "S"  },
  { code: "10 21 00", item: "TOILET COMPARTMENTS",                    type: "S"  },
  { code: "10 28 00", item: "TOILET ACCESSORIES",                     type: "S"  },
  { code: "10 44 00", item: "FIRE PROTECTION SPECIALTIES / EXTINGUISHERS", type: "S" },
  { code: "10 51 00", item: "LOCKERS",                                type: "S"  },
  { code: "10 75 00", item: "FLAGPOLES",                              type: "S"  },
  // ── Division 11 — Equipment ───────────────────────────────────────────────
  { code: "11 00 00", item: "EQUIPMENT (MISC)",                       type: "S"  },
  { code: "11 24 00", item: "MEDICAL EQUIPMENT",                      type: "S"  },
  { code: "11 31 00", item: "RESIDENTIAL APPLIANCES",                 type: "S"  },
  { code: "11 40 00", item: "FOODSERVICE EQUIPMENT",                  type: "S"  },
  // ── Division 12 — Furnishings ─────────────────────────────────────────────
  { code: "12 00 00", item: "FURNISHINGS (MISC)",                     type: "S"  },
  { code: "12 32 00", item: "MANUFACTURED CASEWORK",                  type: "S"  },
  // ── Division 13 — Special Construction ───────────────────────────────────
  { code: "13 00 00", item: "SPECIAL CONSTRUCTION",                   type: "S"  },
  // ── Division 14 — Conveying Equipment ────────────────────────────────────
  { code: "14 20 00", item: "ELEVATORS",                              type: "S"  },
  { code: "14 24 00", item: "HYDRAULIC ELEVATORS",                    type: "S"  },
  // ── Division 21 — Fire Suppression ───────────────────────────────────────
  { code: "21 00 00", item: "FIRE SUPPRESSION",                       type: "S"  },
  // ── Division 22 — Plumbing ───────────────────────────────────────────────
  { code: "22 00 00", item: "PLUMBING",                               type: "S"  },
  // ── Division 23 — HVAC ───────────────────────────────────────────────────
  { code: "23 00 00", item: "HVAC",                                   type: "S"  },
  // ── Division 26 — Electrical ─────────────────────────────────────────────
  { code: "26 00 00", item: "ELECTRICAL",                             type: "S"  },
  // ── Division 27 — Communications ─────────────────────────────────────────
  { code: "27 00 00", item: "COMMUNICATIONS / DATA & VOICE",          type: "S"  },
  // ── Division 28 — Electronic Safety & Security ───────────────────────────
  { code: "28 00 00", item: "ELECTRONIC SAFETY & SECURITY",           type: "S"  },
  { code: "28 13 00", item: "ACCESS CONTROL",                         type: "S"  },
  // ── Division 31 — Earthwork ───────────────────────────────────────────────
  { code: "31 20 00", item: "EARTHWORK / EXCAVATION",                 type: "S"  },
  { code: "31 20 00", item: "UNDERCUT & REPLACE",                     type: "S"  },
  { code: "31 20 00", item: "ROCK EXCAVATION",                        type: "S"  },
  { code: "31 23 19", item: "DEWATERING SYSTEM",                      type: "S"  },
  { code: "31 25 00", item: "EROSION & SEDIMENTATION CONTROLS",       type: "S"  },
  { code: "31 41 00", item: "SHORING",                                type: "S"  },
  { code: "31 63 00", item: "DRILLED CONCRETE PIERS & SHAFTS",        type: "S"  },
  // ── Division 32 — Exterior Improvements ──────────────────────────────────
  { code: "32 11 00", item: "BASE COURSES / SUBBASE",                 type: "S"  },
  { code: "32 12 00", item: "ASPHALT PAVING",                         type: "S"  },
  { code: "32 13 00", item: "CONCRETE PAVING / SITE CONCRETE",        type: "S"  },
  { code: "32 14 00", item: "UNIT PAVING",                            type: "S"  },
  { code: "32 17 00", item: "PARKING LOT STRIPING & PAVEMENT MARKING",type: "S"  },
  { code: "32 31 00", item: "FENCES & GATES",                         type: "S"  },
  { code: "32 71 00", item: "IRRIGATION",                             type: "S"  },
  { code: "32 80 00", item: "LANDSCAPING",                            type: "S"  },
  // ── Division 33 — Utilities ───────────────────────────────────────────────
  { code: "33 10 00", item: "WATER UTILITIES / DOMESTIC WATER",       type: "S"  },
  { code: "33 12 00", item: "WATER DISTRIBUTION / FIRE LOOP & HYDRANTS", type: "S" },
  { code: "33 30 00", item: "SANITARY SEWERAGE",                      type: "S"  },
  { code: "33 40 00", item: "STORM DRAINAGE",                         type: "S"  },
];

const SHEET_2500_ROWS = [
  { section: "SUPERVISION AND ADMINISTRATION", isSection: true },
  { code: "", desc: "FROM SUPERVISION WORKSHEET", unit: "", autoCalc: "supervision", readOnly: true, comments: "Includes burden." },
  { section: "GENERAL REQUIREMENTS", isSection: true },
  { code: "", desc: "WATCHMAN",                    unit: "WK",   sub_up: 0,       comments: "" },
  { code: "", desc: "TEMPORARY FENCE",             unit: "LF",   sub_up: 10,      comments: "" },
  { code: "", desc: "TEMPORARY FENCE - GATES (PAIR)", unit: "EA", sub_up: 500,    comments: "" },
  { code: "", desc: "TEMPORARY ROADS",             unit: "SY",   sub_up: 38,      comments: "PUT DOWN AND TAKE UP" },
  { code: "", desc: "BUILDING CLEANOUT",           unit: "WK",   sub_up: 1000,    comments: "" },
  { code: "", desc: "FINAL CLEANING",              unit: "SF",   sub_up: 0.60,    comments: "" },
  { code: "", desc: "WASTE REMOVAL SERVICE",       unit: "PULL", sub_up: 850,     comments: "USE FOR CONC. WASHOUT" },
  { code: "", desc: "ROUGH HARDWARE & NAILS",      unit: "LS",   sub_up: 5000,    comments: "" },
  { code: "", desc: "JOB BUILT SAFETY RAILS",      unit: "LF",   sub_up: 5,       comments: "" },
  { code: "", desc: "JOB SIGN",                    unit: "EA",   sub_up: 750,     comments: "" },
  { code: "", desc: "TEMPORARY ENCLOSURES",        unit: "LS",   sub_up: 3000,    comments: "" },
  { section: "TEMPORARY FACILITIES", isSection: true },
  { code: "", desc: "ICE AND CUPS",                unit: "MO",   sub_up: 100,     comments: "" },
  { code: "", desc: "OFFICE TRAILER & FURNITURE",  unit: "MO",   sub_up: 1000,    comments: "" },
  { code: "", desc: "SETUP/REMOVE OFFICE TRAILERS",unit: "EA",   sub_up: 3000,    comments: "" },
  { code: "", desc: "TEMPORARY SHEDS/TRAILERS",    unit: "MO",   sub_up: 292,     comments: "" },
  { code: "", desc: "TEMPORARY TOILETS",           unit: "MO-EA",sub_up: 145,     comments: "Unit: # of Months x # of Toilets" },
  { code: "", desc: "TEMPORARY ELECTRIC - JOB OFFICE", unit: "MO", sub_up: 150,  comments: "" },
  { code: "", desc: "TEMPORARY ELECTRIC - PRE MECH START", unit: "MO", sub_up: 0, comments: "Use $0.03/SF/MO" },
  { code: "", desc: "TEMPORARY HEAT",              unit: "MO",   sub_up: 0,       comments: "" },
  { code: "", desc: "TEMPORARY WATER",             unit: "MO",   sub_up: 250,     comments: "" },
  { section: "LIVING & TRAVEL EXPENSES", isSection: true },
  { code: "", desc: "LIVING EXPENSES FROM SUPERVISION WORKSHEET", unit: "", autoCalc: "perDiem", readOnly: true, comments: "Per-diem from 2500-Supervision" },
  { code: "", desc: "TRAVEL EXPENSES FROM SUPERVISION WORKSHEET", unit: "", autoCalc: "travel",  readOnly: true, comments: "Total Travel from 2500-Supervision" },
  { section: "OFFICE EQUIPMENT & SUPPLIES", isSection: true },
  { code: "", desc: "JOB OFFICE EXPENSES",         unit: "MO",   sub_up: 150,     comments: "" },
  { section: "PLAN REPRODUCTION & PHOTOGRAPHS", isSection: true },
  { code: "", desc: "PLANS & SPECIFICATIONS",      unit: "LS",   sub_up: 250,     comments: "" },
  { code: "", desc: "PHOTOGRAPHS",                 unit: "MO",   sub_up: 650,     comments: "" },
  { section: "SAFETY", isSection: true },
  { code: "", desc: "SAFETY DIRECTOR - WAGES & TRAVEL", unit: "", autoCalc: "safetyDir", readOnly: true, comments: "From 2500-Supervision" },
  { code: "", desc: "GENERAL SAFETY SUPPLIES",     unit: "MO",   sub_up: 100,     comments: "" },
  { code: "", desc: "DRUG SCREENING/TESTING",      unit: "MO",   sub_up: 100,     comments: "" },
];

const SHEET_2600_ROWS = [
  { section: "EQUIPMENT & TRANSPORTATION", isSection: true },
  { code: "2600-080", desc: "EQUIPMENT RENTAL - COMPANY OWNED (FROM SUMMARY BELOW)", unit: "MO",  sub_up: 0,    comments: "See Equipment Rental Summary section" },
  { code: "2600-081", desc: "EQUIPMENT RENTAL - FROM OTHERS (FROM SUMMARY BELOW)",   unit: "MO",  sub_up: 0,    comments: "See Equipment Rental Summary section" },
  { code: "",         desc: "FUEL - COMPANY TRUCKS - FROM WORKSHEET",                unit: "",    autoCalc: "fuel", readOnly: true, comments: "See 2500-Supervision" },
  { code: "2600-100", desc: "GAS & OIL - OTHER EQUIPMENT",      unit: "MO",  sub_up: 250,  comments: "" },
  { code: "2600-110", desc: "SMALL TOOLS",                      unit: "LS",  sub_up: 0,    comments: "" },
  { code: "2600-120", desc: "TOOLS IN AND OUT",                 unit: "MO",  sub_up: 0,    comments: "" },
  { section: "HOISTING", isSection: true },
  { code: "2600-360", desc: "PERSONNEL HOIST RENTAL",           unit: "MO",  sub_up: 0,    comments: "" },
  { code: "2600-362", desc: "HOIST FREIGHT",                    unit: "LS",  sub_up: 0,    comments: "" },
  { code: "2600-364", desc: "HOIST UP-DOWN-MOVE",               unit: "LS",  sub_up: 0,    comments: "" },
  { code: "2600-366", desc: "HOIST FOUNDATION & DOCK",          unit: "LS",  sub_up: 0,    comments: "" },
  { code: "2600-368", desc: "PERSONNEL HOIST OPERATOR",         unit: "WK",  sub_up: 0,    comments: "" },
  { code: "2600-390", desc: "CRANE RENTAL",                     unit: "MO",  sub_up: 0,    comments: "" },
  { code: "2600-392", desc: "CRANE OPERATOR",                   unit: "WK",  sub_up: 0,    comments: "" },
  { section: "EQUIPMENT RENTAL SUMMARY — COMPANY OWNED", isSection: true },
  { code: "",  desc: "PICKUP TRUCKS FROM WORKSHEET",            unit: "",    autoCalc: "fuel", readOnly: true, comments: "See 2500-Supervision" },
  { code: "",  desc: "LEVELS",                                  unit: "MO",  sub_up: 306,  comments: "" },
  { code: "",  desc: "TRANSITS",                                unit: "MO",  sub_up: 312,  comments: "" },
  { code: "",  desc: "LASERS",                                  unit: "MO",  sub_up: 387,  comments: "" },
  { code: "",  desc: "TOTAL STATIONS",                          unit: "MO",  sub_up: 1060, comments: "" },
  { code: "",  desc: "LASER SCANNER",                           unit: "USE", sub_up: 1700, comments: "" },
  { code: "",  desc: "DRONES",                                  unit: "USE", sub_up: 100,  comments: "" },
  { code: "",  desc: "IPADS",                                   unit: "MO",  sub_up: 200,  comments: "" },
  { code: "",  desc: "COMPUTERS FROM WORKSHEET",                unit: "",    autoCalc: "fuel", readOnly: true, comments: "See 2500-Supervision" },
  { code: "",  desc: "BIM360",                                  unit: "$-TOTAL", sub_up: 0, comments: "Plug total project cost into QTY; rate = 0.0025" },
  { code: "",  desc: "CELL PHONES FROM WORKSHEET",              unit: "",    autoCalc: "fuel", readOnly: true, comments: "See 2500-Supervision" },
  { code: "",  desc: "OTHER FROM WORKSHEET",                    unit: "",    autoCalc: "fuel", readOnly: true, comments: "See 2500-Supervision" },
  { code: "",  desc: "WI-FI HOT SPOT",                         unit: "MO",  sub_up: 100,  comments: "" },
  { code: "",  desc: "BACKHOE",                                 unit: "MO",  sub_up: 1351, comments: "" },
  { code: "",  desc: "ATV (KUBOTA)",                            unit: "MO",  sub_up: 645,  comments: "" },
  { code: "",  desc: "16' UTILITY TRAILER",                     unit: "MO",  sub_up: 301,  comments: "" },
  { code: "",  desc: "AIR COMPRESSOR IR",                       unit: "MO",  sub_up: 48,   comments: "" },
  { section: "SAWS", isSection: true },
  { code: "",  desc: "SKILL SAW",                               unit: "MO",  sub_up: 52,   comments: "12 Mo. Cap" },
  { code: "",  desc: "SAWZALL AC",                              unit: "MO",  sub_up: 43,   comments: "12 Mo. Cap" },
  { code: "",  desc: "SAWZALL DC",                              unit: "MO",  sub_up: 43,   comments: "12 Mo. Cap" },
  { code: "",  desc: "JIG SAWS",                                unit: "MO",  sub_up: 28,   comments: "12 Mo. Cap" },
  { code: "",  desc: "ROTO ZIP SAWS",                           unit: "MO",  sub_up: 28,   comments: "12 Mo. Cap" },
  { code: "",  desc: "QUICKIE SAW",                             unit: "MO",  sub_up: 181,  comments: "" },
  { code: "",  desc: "HYDRO PACK SAW",                          unit: "MO",  sub_up: 469,  comments: "" },
  { code: "",  desc: "PORTA BAND SAW",                          unit: "MO",  sub_up: 60,   comments: "12 Mo. Cap" },
  { code: "",  desc: "CHOP SAWS",                               unit: "MO",  sub_up: 38,   comments: "12 Mo. Cap" },
  { code: "",  desc: "COMPOUND MITER SAW",                      unit: "MO",  sub_up: 117,  comments: "12 Mo. Cap" },
  { code: "",  desc: "TABLE SAW",                               unit: "MO",  sub_up: 78,   comments: "12 Mo. Cap" },
  { section: "DRILLS & DRIVERS", isSection: true },
  { code: "",  desc: "DRILLS DC",                               unit: "MO",  sub_up: 59,   comments: "12 Mo. Cap" },
  { code: "",  desc: "DRILLS AC",                               unit: "MO",  sub_up: 43,   comments: "12 Mo. Cap" },
  { code: "",  desc: "SCREWGUNS DC",                            unit: "MO",  sub_up: 37,   comments: "12 Mo. Cap" },
  { code: "",  desc: "SCREWGUNS AC",                            unit: "MO",  sub_up: 20,   comments: "12 Mo. Cap" },
  { code: "",  desc: "HAMMER DRILLS",                           unit: "MO",  sub_up: 108,  comments: "12 Mo. Cap" },
  { code: "",  desc: "CORE DRILL",                              unit: "MO",  sub_up: 315,  comments: "" },
  { code: "",  desc: "IMPACT WRENCH",                           unit: "MO",  sub_up: 62,   comments: "12 Mo. Cap" },
  { code: "",  desc: "DC COMBO SETS",                           unit: "MO",  sub_up: 128,  comments: "12 Mo. Cap" },
  { section: "DEMOLITION TOOLS", isSection: true },
  { code: "",  desc: "DEMO HAMMER/ELEC",                        unit: "MO",  sub_up: 88,   comments: "12 Mo. Cap" },
  { code: "",  desc: "JACK HAMMER 90# AIR",                     unit: "MO",  sub_up: 144,  comments: "12 Mo. Cap" },
  { code: "",  desc: "CHIPPING HAMMER AIR",                     unit: "MO",  sub_up: 88,   comments: "12 Mo. Cap" },
  { section: "OFFICE EQUIPMENT", isSection: true },
  { code: "",  desc: "COPIER - COMBO NEW STYLE",                unit: "MO",  sub_up: 410,  comments: "18 Mo. Cap" },
  { code: "",  desc: "COPIER dp190-8000",                       unit: "MO",  sub_up: 258,  comments: "18 Mo. Cap" },
  { section: "GRINDERS & SANDERS", isSection: true },
  { code: "",  desc: "GRINDER 4-1/2 INCH",                      unit: "MO",  sub_up: 35,   comments: "12 Mo. Cap" },
  { code: "",  desc: "BELT SANDER",                             unit: "MO",  sub_up: 45,   comments: "12 Mo. Cap" },
  { code: "",  desc: "ROUTER",                                  unit: "MO",  sub_up: 43,   comments: "12 Mo. Cap" },
  { code: "",  desc: "PALM SANDERS",                            unit: "MO",  sub_up: 28,   comments: "12 Mo. Cap" },
  { code: "",  desc: "PLANER",                                  unit: "MO",  sub_up: 27,   comments: "12 Mo. Cap" },
  { section: "SPECIALTY TOOLS", isSection: true },
  { code: "",  desc: "POWDER ACTIVATED TOOL",                   unit: "MO",  sub_up: 96,   comments: "12 Mo. Cap" },
  { code: "",  desc: "TORCH SETS/CART",                         unit: "MO",  sub_up: 72,   comments: "12 Mo. Cap" },
  { code: "",  desc: "WALKIE TALKIE",                           unit: "MO",  sub_up: 35,   comments: "12 Mo. Cap" },
  { code: "",  desc: "GUN SAFE",                                unit: "MO",  sub_up: 155,  comments: "12 Mo. Cap" },
  { section: "HEATING, COOLING & VENTILATION", isSection: true },
  { code: "",  desc: "HEATERS 375,000 BTU",                     unit: "MO",  sub_up: 88,   comments: "12 Mo. Cap" },
  { code: "",  desc: "PORTABLE HEATERS - 1,250,000 BTU",        unit: "MO",  sub_up: 344,  comments: "" },
  { code: "",  desc: "SMALL DEHUMIDIFIER",                      unit: "MO",  sub_up: 43,   comments: "" },
  { code: "",  desc: "LARGE DEHUMIDIFIER",                      unit: "MO",  sub_up: 303,  comments: "" },
  { code: "",  desc: "SHOP/BOX FANS",                           unit: "MO",  sub_up: 51,   comments: "12 Mo. Cap" },
  { section: "MATERIAL HANDLING", isSection: true },
  { code: "",  desc: "GANG BOX",                                unit: "MO",  sub_up: 90,   comments: "12 Mo. Cap" },
  { code: "",  desc: "TRASH CARTS",                             unit: "MO",  sub_up: 86,   comments: "12 Mo. Cap" },
  { code: "",  desc: "2 1/2 YD HOPPER",                        unit: "MO",  sub_up: 284,  comments: "" },
  { code: "",  desc: "1 YD HOPPER",                             unit: "MO",  sub_up: 163,  comments: "" },
  { code: "",  desc: "PALLET JACK",                             unit: "MO",  sub_up: 48,   comments: "12 Mo. Cap" },
  { code: "",  desc: "ROCK DOLLY",                              unit: "MO",  sub_up: 43,   comments: "12 Mo. Cap" },
  { section: "POWER & PUMPS", isSection: true },
  { code: "",  desc: "HONDA GENERATORS",                        unit: "MO",  sub_up: 292,  comments: "12 Mo. Cap" },
  { code: "",  desc: "AIR COMPRESSOR / GAS",                    unit: "MO",  sub_up: 144,  comments: "12 Mo. Cap" },
  { code: "",  desc: "AIR COMPRESSOR / ELEC",                   unit: "MO",  sub_up: 104,  comments: "12 Mo. Cap" },
  { code: "",  desc: "SUMP PUMPS / ELEC",                       unit: "MO",  sub_up: 43,   comments: "12 Mo. Cap" },
  { code: "",  desc: "TRASH PUMPS / GAS",                       unit: "MO",  sub_up: 103,  comments: "12 Mo. Cap" },
  { section: "WELDING & FINISHING", isSection: true },
  { code: "",  desc: "WELDERS/BOBCAT/GAS",                      unit: "MO",  sub_up: 440,  comments: "12 Mo. Cap" },
  { code: "",  desc: "WELDERS WIRE",                            unit: "MO",  sub_up: 88,   comments: "12 Mo. Cap" },
  { code: "",  desc: "AIR NAILERS",                             unit: "MO",  sub_up: 48,   comments: "12 Mo. Cap" },
  { section: "VACUUMS & CLEANING", isSection: true },
  { code: "",  desc: "UPRIGHT VAC",                             unit: "MO",  sub_up: 66,   comments: "12 Mo. Cap" },
  { code: "",  desc: "55 GAL VAC",                              unit: "MO",  sub_up: 199,  comments: "" },
  { code: "",  desc: "SHOP VAC",                                unit: "MO",  sub_up: 31,   comments: "12 Mo. Cap" },
  { code: "",  desc: "WEED EATER",                              unit: "MO",  sub_up: 56,   comments: "12 Mo. Cap" },
  { code: "",  desc: "PRESSURE WASHER",                         unit: "MO",  sub_up: 172,  comments: "12 Mo. Cap" },
  { section: "ICRA — INFECTION CONTROL (HEALTHCARE)", isSection: true },
  { code: "",  desc: "NEG. AIR MACHINES",                       unit: "MO",  sub_up: 181,  comments: "12 Mo. Cap" },
  { code: "",  desc: "NEG. AIR SCRUBBER",                       unit: "MO",  sub_up: 320,  comments: "" },
  { code: "",  desc: "PORTABLE PRESSURE MONITOR",               unit: "MO",  sub_up: 171,  comments: "" },
  { code: "",  desc: "DUST CONTAINMENT KIT",                    unit: "EA",  sub_up: 845,  comments: "" },
  { code: "",  desc: "DOOR PANEL",                              unit: "MO",  sub_up: 215,  comments: "Ante Room Enclosure" },
  { code: "",  desc: "48-INCH PANEL",                              unit: "MO",  sub_up: 129,  comments: "Ante Room Enclosure" },
  { code: "",  desc: "24-INCH PANEL",                              unit: "MO",  sub_up: 97,   comments: "Ante Room Enclosure" },
  { code: "",  desc: "24-INCH PANEL W/ HEPA PORT",                 unit: "MO",  sub_up: 112,  comments: "Ante Room Enclosure" },
  { code: "",  desc: "PRESS GAGE KIT",                          unit: "MO",  sub_up: 66,   comments: "" },
  { code: "",  desc: "HINGED CORNER PANEL",                     unit: "MO",  sub_up: 47,   comments: "" },
  { code: "",  desc: "CLOSURE STRIPS",                          unit: "EA",  sub_up: 19,   comments: "Pkg of 12" },
  { code: "",  desc: "PANEL CART",                              unit: "MO",  sub_up: 57,   comments: "" },
  { code: "",  desc: "PANEL EXTENSION - 24-INCH",                  unit: "MO",  sub_up: 18,   comments: "" },
  { section: "EQUIPMENT RENTAL — FROM OTHERS", isSection: true },
  { code: "-", desc: "TRACKHOE",                                unit: "MO",  sub_up: 6000, comments: "" },
  { code: "-", desc: "BACKHOE",                                 unit: "MO",  sub_up: 1850, comments: "" },
  { code: "-", desc: "DOZER",                                   unit: "MO",  sub_up: 3200, comments: "" },
  { code: "-", desc: "TRUCK RENTAL",                            unit: "MO",  sub_up: 2000, comments: "" },
  { code: "-", desc: "EQUIPMENT RENTAL EXCAV.",                 unit: "LS",  sub_up: 5400, comments: "" },
  { code: "-", desc: "BUY ELECTRIC WATER PUMP & HOSE",         unit: "EA",  sub_up: 500,  comments: "" },
  { code: "-", desc: "CONCRETE BUCKETS",                        unit: "MO",  sub_up: 210,  comments: "" },
  { code: "-", desc: "FINISHING MACHINE (4 UNITS)",             unit: "MO",  sub_up: 1260, comments: "4 @ $315/Mo" },
  { code: "-", desc: "375 CFM AIR COMPRESSOR",                  unit: "MO",  sub_up: 1200, comments: "" },
];

const SHEET_2700_ROWS = [
  { section: "PERMITS & INSURANCE", isSection: true },
  { code: "2700-240", desc: "PERMITS & INSPECTION FEES",     unit: "SF",  sub_up: 0,    comments: "" },
  { code: "2700-290", desc: "BUILDER'S RISK INSURANCE",      unit: "MO",  sub_up: 94,   comments: "" },
  { code: "2700-295", desc: "ALL RISK INSURANCE OR D.I.C.",  unit: "$$",  sub_up: 0,    comments: "" },
  { code: "2700-300", desc: "OWNER'S PROTECTIVE",            unit: "$$",  sub_up: 0,    comments: "" },
  { section: "SUBCONTRACTOR BONDS", isSection: true },
  { code: "2700-400", desc: "SUBCONTRACTOR BONDS",           unit: "LS",  sub_up: 0,    comments: "" },
  { code: "2700-900", desc: "SUBGUARD",                      unit: "LS",  sub_up: 0,    comments: "" },
  { section: "SALES & USE TAX", isSection: true },
  { code: "2800-200", desc: "SALES TAX FOR ENTERPRISE ZONE", unit: "LS",  sub_up: 0,    comments: "" },
];

const SHEET_2900_ROWS = [
  { section: "TESTING & INSPECTION", isSection: true },
  { code: "2900-100", desc: "TESTING AND INSPECTION",        unit: "LS",  sub_up: 0,    comments: "" },
  { code: "",         desc: "COMPACTION TESTING",            unit: "LS",  sub_up: 0,    comments: "" },
  { code: "",         desc: "SOILS ENGR FOR PILE LOAD TEST", unit: "LS",  sub_up: 0,    comments: "" },
  { code: "",         desc: "PILE INSPECTION & LOGS",        unit: "LS",  sub_up: 0,    comments: "" },
  { code: "",         desc: "DRILL PIER INSPECTION",         unit: "LS",  sub_up: 0,    comments: "" },
  { code: "",         desc: "CONCRETE TESTING",              unit: "LS",  sub_up: 0,    comments: "" },
  { code: "",         desc: "STEEL TESTING & INSPECTION",    unit: "LS",  sub_up: 0,    comments: "" },
  { code: "",         desc: "FIREPROOFING TESTING & INSPECTION", unit: "LS", sub_up: 0, comments: "" },
  { section: "CONTRACT ALLOWANCES", isSection: true },
  { code: "2900-200", desc: "",  unit: "LS",  sub_up: 0, comments: "" },
  { code: "2900-201", desc: "",  unit: "LS",  sub_up: 0, comments: "" },
  { code: "2900-202", desc: "",  unit: "LS",  sub_up: 0, comments: "" },
  { code: "2900-203", desc: "",  unit: "LS",  sub_up: 0, comments: "" },
  { code: "2900-204", desc: "",  unit: "LS",  sub_up: 0, comments: "" },
];

// ── helpers ───────────────────────────────────────────────────────────────────
const fmtDollar = (v) => {
  if (!v && v !== 0) return "—";
  return "$" + Number(v).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
};
const fmtDollarInput = (v) => (v === "" || v === null || v === undefined) ? "" : String(v);
const parseDollar = (s) => parseFloat(s) || 0;

// ── Shared detail-sheet table ─────────────────────────────────────────────────
function DetailSheet({ title, code, rows: templateRows, sheetData, onSheetChange, autoCalcValues }) {
  // sheetData is { [rowIdx]: { qty, sub_up, sub_total, labor_up, labor, mat_up, material, comments } }

  function getRow(idx) {
    return sheetData[idx] || {};
  }
  function setRowField(idx, field, value) {
    const existing = sheetData[idx] || {};
    const updated  = { ...existing, [field]: value };
    // auto-calc sub_total
    const qty    = parseDollar(field === "qty"    ? value : (updated.qty    ?? templateRows[idx]?.qty    ?? 0));
    const sub_up = parseDollar(field === "sub_up" ? value : (updated.sub_up ?? templateRows[idx]?.sub_up ?? 0));
    updated.sub_total = qty * sub_up;
    onSheetChange({ ...sheetData, [idx]: updated });
  }

  const totalSub = templateRows.reduce((acc, row, idx) => {
    if (row.isSection) return acc;
    if (row.readOnly && row.autoCalc) {
      return acc + (autoCalcValues?.[row.autoCalc] || 0);
    }
    const d = getRow(idx);
    const qty    = parseDollar(d.qty    ?? 0);
    const sub_up = parseDollar(d.sub_up ?? row.sub_up ?? 0);
    return acc + qty * sub_up;
  }, 0);

  const thStyle = {
    padding: "6px 10px", fontSize: 11, fontWeight: 700,
    color: C.textMuted, textTransform: "uppercase", letterSpacing: "0.06em",
    borderBottom: `2px solid ${C.border}`, textAlign: "left",
    whiteSpace: "nowrap", backgroundColor: C.contentBg, position: "sticky", top: 0, zIndex: 2,
  };
  const tdStyle = { padding: "5px 10px", fontSize: 12, borderBottom: `1px solid ${C.border}`, color: C.textPrimary };
  const inputStyle = {
    width: "100%", border: `1px solid ${C.border}`, borderRadius: 3,
    padding: "3px 6px", fontSize: 12, fontFamily: "monospace",
    backgroundColor: C.white, color: C.textPrimary, outline: "none",
    boxSizing: "border-box",
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100%", overflow: "hidden" }}>
      {/* Sheet header */}
      <div style={{ padding: "12px 20px 10px", borderBottom: `1px solid ${C.border}`, flexShrink: 0, display: "flex", alignItems: "baseline", gap: 12 }}>
        <span style={{ fontSize: 11, fontWeight: 700, color: C.textMuted, fontFamily: "monospace" }}>{code}</span>
        <span style={{ fontSize: 15, fontWeight: 700, color: C.textPrimary }}>{title}</span>
        <span style={{ marginLeft: "auto", fontSize: 12, color: C.textMuted }}>
          TOTAL: <strong style={{ color: C.textPrimary, fontFamily: "monospace" }}>{fmtDollar(totalSub)}</strong>
        </span>
      </div>

      <div style={{ flex: 1, overflowY: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead>
            <tr>
              <th style={{ ...thStyle, width: 90  }}>Code</th>
              <th style={{ ...thStyle            }}>Description</th>
              <th style={{ ...thStyle, width: 60  }}>Unit</th>
              <th style={{ ...thStyle, width: 80, textAlign: "right" }}>Qty</th>
              <th style={{ ...thStyle, width: 100, textAlign: "right" }}>Unit Price</th>
              <th style={{ ...thStyle, width: 120, textAlign: "right" }}>Sub/Equip $</th>
              <th style={{ ...thStyle, width: 200 }}>Comments</th>
            </tr>
          </thead>
          <tbody>
            {templateRows.map((row, idx) => {
              if (row.isSection) return (
                <tr key={idx}>
                  <td colSpan={7} style={{
                    padding: "10px 10px 4px", fontSize: 11, fontWeight: 700,
                    color: C.textMuted, textTransform: "uppercase", letterSpacing: "0.08em",
                    backgroundColor: C.contentBg, borderBottom: `1px solid ${C.border}`,
                  }}>{row.section}</td>
                </tr>
              );

              const d = getRow(idx);
              const isAutoCalc = row.readOnly && row.autoCalc;
              const autoVal = isAutoCalc ? (autoCalcValues?.[row.autoCalc] || 0) : null;
              const qty    = isAutoCalc ? null : (d.qty    ?? "");
              const sub_up = isAutoCalc ? null : (d.sub_up !== undefined ? d.sub_up : (row.sub_up !== undefined ? row.sub_up : ""));
              const subTotal = isAutoCalc ? autoVal : parseDollar(qty) * parseDollar(sub_up);

              return (
                <tr key={idx} style={{ backgroundColor: idx % 2 === 0 ? C.white : C.tableStripe }}>
                  <td style={{ ...tdStyle, fontFamily: "monospace", fontSize: 11, color: C.textMuted }}>{row.code || ""}</td>
                  <td style={{ ...tdStyle, color: isAutoCalc ? C.textMuted : C.textPrimary, fontStyle: isAutoCalc ? "italic" : "normal" }}>
                    {row.desc === "" && !isAutoCalc ? (
                      <input
                        style={{ ...inputStyle, fontFamily: "Inter, sans-serif", width: "100%" }}
                        type="text"
                        value={d.desc ?? ""}
                        onChange={e => setRowField(idx, "desc", e.target.value)}
                        placeholder="Enter allowance description…"
                      />
                    ) : row.desc}
                  </td>
                  <td style={{ ...tdStyle, color: C.textMuted, fontSize: 11 }}>{row.unit}</td>
                  <td style={{ ...tdStyle, textAlign: "right" }}>
                    {!isAutoCalc && (
                      <input
                        style={inputStyle}
                        type="number" min="0" step="any"
                        value={qty}
                        onChange={e => setRowField(idx, "qty", e.target.value)}
                        placeholder="0"
                      />
                    )}
                  </td>
                  <td style={{ ...tdStyle, textAlign: "right" }}>
                    {!isAutoCalc && (
                      <input
                        style={inputStyle}
                        type="number" min="0" step="0.01"
                        value={sub_up !== "" ? sub_up : ""}
                        onChange={e => setRowField(idx, "sub_up", e.target.value)}
                        placeholder="0.00"
                      />
                    )}
                  </td>
                  <td style={{ ...tdStyle, textAlign: "right", fontFamily: "monospace", fontWeight: subTotal > 0 ? 600 : 400, color: subTotal > 0 ? C.textPrimary : C.textLight }}>
                    {subTotal > 0 ? fmtDollar(subTotal) : "—"}
                  </td>
                  <td style={{ ...tdStyle }}>
                    <input
                      style={{ ...inputStyle, fontFamily: "Inter, sans-serif" }}
                      type="text"
                      value={d.comments ?? row.comments ?? ""}
                      onChange={e => setRowField(idx, "comments", e.target.value)}
                      placeholder=""
                    />
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      {/* Totals footer */}
      <div style={{ padding: "10px 20px", borderTop: `2px solid ${C.border}`, backgroundColor: C.contentBg, display: "flex", justifyContent: "flex-end", gap: 40, flexShrink: 0 }}>
        <div style={{ textAlign: "right" }}>
          <div style={{ fontSize: 11, color: C.textMuted, textTransform: "uppercase", letterSpacing: "0.06em" }}>Total {title}</div>
          <div style={{ fontSize: 16, fontWeight: 700, fontFamily: "monospace", color: C.textPrimary }}>{fmtDollar(totalSub)}</div>
        </div>
      </div>
    </div>
  );
}

// ── 2500-Supervision Sheet ────────────────────────────────────────────────────
function SupervisionSheet({ supervisionData, onSupervisionChange, durationWeeks, firmRates, insSsRate, onInsSsRateChange, assumptions: assumptionsProp }) {

  // Assumptions come from parent (EstimateWorkspace via Settings modal)
  const assumptions = assumptionsProp || SUPERVISION_ASSUMPTION_DEFAULTS;

  function getRole(key) { return supervisionData[key] || {}; }
  function setRoleField(key, field, value) {
    onSupervisionChange({ ...supervisionData, [key]: { ...getRole(key), [field]: value } });
  }

  // ── Per-role calculations matching the Excel logic ────────────────────────
  function calcRole(key) {
    const d    = getRole(key);
    const role = SUPERVISION_ROLES.find(r => r.key === key);
    const A    = assumptions;
    const hourlyRate = firmRates[key] ?? role?.defaultRate ?? 0;

    // Salary
    const weeks  = parseDollar(d.weeks ?? 0);
    const salary = weeks * hourlyRate * A.hoursPerWeek;

    // Fleet vehicle fuel
    const hasFleet      = d.fleetVehicle === "Yes";
    const dailyMiles    = parseDollar(d.dailyMiles ?? A.defaultDailyMiles);
    const fleetFuel     = hasFleet ? (weeks * 5 * dailyMiles / A.mpg * A.fuelPerGal) : 0;

    // Computer, cell, other (monthly rates × months)
    const months        = weeks / A.weeksPerMonth;
    const computerAmt   = parseDollar(d.computerMo ?? 0) * months;
    const cellAmt       = parseDollar(d.cellMo     ?? 0) * months;
    const otherEquipAmt = parseDollar(d.otherEquipMo ?? 0) * months;

    // Per-diem
    const hasPerDiem    = d.perDiemYN === "Yes";
    const perDiemMoAmt  = parseDollar(d.perDiemMo ?? 0);
    const perDiemTotal  = hasPerDiem ? perDiemMoAmt * months : 0;

    // Personal vehicle mileage trips
    const mileageTrips  = parseDollar(d.mileageTrips   ?? 0);
    const milesPerTrip  = parseDollar(d.milesPerTrip   ?? 155);
    const mileageCost   = mileageTrips * milesPerTrip * A.mileageRate;

    // Other travel (trips with hotel/plane/car)
    const travelTrips   = parseDollar(d.travelTrips    ?? 0);
    const daysPerTrip   = parseDollar(d.daysPerTrip    ?? 0);
    const hotelRequired = d.hotelRequired === "Yes";
    const planeTicket   = d.planeTicket   === "Yes";
    const carRental     = d.carRental     === "Yes";
    const carMiPerTrip  = parseDollar(d.carMiPerTrip   ?? 0);

    const hotelCost   = hotelRequired ? travelTrips * daysPerTrip * A.hotelRate    : 0;
    const planeCost   = planeTicket   ? travelTrips * A.planeRate                  : 0;
    const carCost     = carRental     ? travelTrips * daysPerTrip * A.carRentalRate: 0;
    const carMileCost = carMiPerTrip  > 0 ? travelTrips * carMiPerTrip * A.mileageRate : 0;
    const mealsCost   = travelTrips * daysPerTrip * A.foodAllowance;
    const travelTotal = hotelCost + planeCost + carCost + carMileCost + mealsCost;

    const total = salary + fleetFuel + computerAmt + cellAmt + otherEquipAmt + perDiemTotal + mileageCost + travelTotal;
    const loadedHourlyRate = weeks > 0 ? total / (weeks * A.hoursPerWeek) : 0;

    return { salary, fleetFuel, computerAmt, cellAmt, otherEquipAmt,
             perDiemTotal, mileageCost, travelTotal, total, hourlyRate, loadedHourlyRate, months };
  }

  const totals = SUPERVISION_ROLES.reduce((acc, role) => {
    const c = calcRole(role.key);
    acc.salary        += c.salary;
    acc.fleetFuel     += c.fleetFuel;
    acc.equipment     += c.computerAmt + c.cellAmt + c.otherEquipAmt;
    acc.perDiem       += c.perDiemTotal;
    acc.mileage       += c.mileageCost;
    acc.travel        += c.travelTotal;
    acc.total         += c.total;
    return acc;
  }, { salary:0, fleetFuel:0, equipment:0, perDiem:0, mileage:0, travel:0, total:0 });

  // Column group header style
  const grpHdr = (label, cols, color) => (
    <th colSpan={cols} style={{
      padding: "5px 8px", fontSize: 10, fontWeight: 700, textAlign: "center",
      textTransform: "uppercase", letterSpacing: "0.06em",
      backgroundColor: color, color: "#1E293B",
      borderBottom: `1px solid ${C.border}`, borderRight: `2px solid ${C.border}`,
      whiteSpace: "nowrap",
    }}>{label}</th>
  );

  const thS = {
    padding: "5px 8px", fontSize: 10, fontWeight: 700, color: C.textMuted,
    textTransform: "uppercase", letterSpacing: "0.05em",
    borderBottom: `2px solid ${C.border}`, textAlign: "right",
    backgroundColor: C.contentBg, whiteSpace: "nowrap",
    position: "sticky", top: 28, zIndex: 1,
  };
  const tdS = { padding: "4px 7px", fontSize: 11, borderBottom: `1px solid ${C.border}`, verticalAlign: "middle" };

  const numIn = (key, field, w, placeholder) => (
    <input
      style={{ width: w||60, border:`1px solid ${C.border}`, borderRadius:3, padding:"2px 4px",
               fontSize:11, fontFamily:"monospace", backgroundColor:C.white, textAlign:"right", boxSizing:"border-box" }}
      type="number" min="0" step="any"
      value={getRole(key)[field] ?? ""}
      onChange={e => setRoleField(key, field, e.target.value)}
      placeholder={placeholder||"0"}
    />
  );

  const ynSel = (key, field) => (
    <select
      style={{ width:52, border:`1px solid ${C.border}`, borderRadius:3, padding:"2px 3px",
               fontSize:11, backgroundColor:C.white, boxSizing:"border-box" }}
      value={getRole(key)[field] ?? "No"}
      onChange={e => setRoleField(key, field, e.target.value)}
    >
      <option>No</option><option>Yes</option>
    </select>
  );

  const moneyCell = (val) => (
    <td style={{ ...tdS, textAlign:"right", fontFamily:"monospace",
                 fontWeight: val > 0 ? 600 : 400,
                 color: val > 0 ? C.textPrimary : C.textLight }}>
      {val > 0 ? fmtDollar(val) : "—"}
    </td>
  );

  const borderRight = { borderRight: `2px solid ${C.border}` };

  return (
    <div style={{ display:"flex", flexDirection:"column", height:"100%", overflow:"hidden" }}>

      {/* Header bar */}
      <div style={{ padding:"10px 20px", borderBottom:`1px solid ${C.border}`, flexShrink:0,
                    display:"flex", alignItems:"center", justifyContent:"space-between" }}>
        <span style={{ fontSize:15, fontWeight:700, color:C.textPrimary }}>Supervision & Administration</span>
        <div style={{ display:"flex", gap:8, alignItems:"center" }}>
          <span style={{ fontSize:12, color:C.textMuted }}>
            Total: <strong style={{ fontFamily:"monospace", color:C.textPrimary }}>{fmtDollar(totals.total)}</strong>
          </span>

        </div>
      </div>

      {/* Scrollable table */}
      <div style={{ flex:1, overflowX:"auto", overflowY:"auto" }}>
        <table style={{ borderCollapse:"collapse", minWidth:1400 }}>
          <thead style={{ position:"sticky", top:0, zIndex:2 }}>
            {/* Group headers */}
            <tr>
              <th style={{ ...thS, top:0, textAlign:"left", width:160, borderRight:`2px solid ${C.border}` }} rowSpan={2}>Role</th>
              {grpHdr("Salary", 3, "#DBEAFE")}
              {grpHdr("Fleet Vehicle & Fuel", 3, "#D1FAE5")}
              {grpHdr("Computers / Phones / Equip", 3, "#EDE9FE")}
              {grpHdr("Per-Diem", 2, "#FEF3C7")}
              {grpHdr("Personal Vehicle Mileage", 3, "#FFE4E6")}
              {grpHdr("Other Travel", 6, "#F0FDF4")}
              {grpHdr("Totals", 2, "#F1F5F9")}
            </tr>
            {/* Sub-headers */}
            <tr>
              <th style={{ ...thS }}>Hrly Rate</th>
              <th style={{ ...thS }}>Weeks</th>
              <th style={{ ...thS, ...borderRight }}>Salary $</th>
              <th style={{ ...thS }}>Fleet Veh.</th>
              <th style={{ ...thS }}>Daily Mi.</th>
              <th style={{ ...thS, ...borderRight }}>Fuel $</th>
              <th style={{ ...thS }}>Computer/Mo</th>
              <th style={{ ...thS }}>Cell/Mo</th>
              <th style={{ ...thS, ...borderRight }}>Other/Mo</th>
              <th style={{ ...thS }}>Per Diem</th>
              <th style={{ ...thS, ...borderRight }}>Per Diem $</th>
              <th style={{ ...thS }}>Trips</th>
              <th style={{ ...thS }}>Mi/Trip</th>
              <th style={{ ...thS, ...borderRight }}>Mileage $</th>
              <th style={{ ...thS }}>Trips</th>
              <th style={{ ...thS }}>Days/Trip</th>
              <th style={{ ...thS }}>Hotel</th>
              <th style={{ ...thS }}>Plane</th>
              <th style={{ ...thS }}>Car Rental</th>
              <th style={{ ...thS, ...borderRight }}>Travel $</th>
              <th style={{ ...thS }}>Total Cost</th>
              <th style={{ ...thS }}>Loaded $/hr</th>
            </tr>
          </thead>
          <tbody>
            {SUPERVISION_ROLES.map((role, i) => {
              const c   = calcRole(role.key);
              const d   = getRole(role.key);
              const bg  = i % 2 === 0 ? C.white : C.tableStripe;
              return (
                <tr key={role.key} style={{ backgroundColor: bg }}>
                  {/* Role name */}
                  <td style={{ ...tdS, fontWeight:500, whiteSpace:"nowrap", ...borderRight, position:"sticky", left:0, backgroundColor:bg, zIndex:1 }}>
                    {role.label}
                  </td>
                  {/* Salary */}
                  <td style={{ ...tdS, textAlign:"right", fontFamily:"monospace", fontSize:11, color:C.textMuted }}>${c.hourlyRate.toFixed(2)}</td>
                  <td style={{ ...tdS, textAlign:"right" }}>{numIn(role.key,"weeks",54,"0")}</td>
                  {moneyCell(c.salary)}
                  {/* Fleet */}
                  <td style={{ ...tdS, textAlign:"center" }}>{ynSel(role.key,"fleetVehicle")}</td>
                  <td style={{ ...tdS, textAlign:"right" }}>{numIn(role.key,"dailyMiles",54,String(assumptions.defaultDailyMiles))}</td>
                  {moneyCell(c.fleetFuel)}
                  {/* Equipment */}
                  <td style={{ ...tdS, textAlign:"right" }}>{numIn(role.key,"computerMo",70,"0")}</td>
                  <td style={{ ...tdS, textAlign:"right" }}>{numIn(role.key,"cellMo",60,"0")}</td>
                  <td style={{ ...tdS, textAlign:"right", ...borderRight }}>{numIn(role.key,"otherEquipMo",60,"0")}</td>
                  {/* Per-Diem */}
                  <td style={{ ...tdS, textAlign:"center" }}>{ynSel(role.key,"perDiemYN")}</td>
                  <td style={{ ...tdS, textAlign:"right", ...borderRight }}>
                    {d.perDiemYN === "Yes" ? numIn(role.key,"perDiemMo",80,"0") : <span style={{ color:C.textLight, fontSize:11 }}>—</span>}
                    {d.perDiemYN === "Yes" && c.perDiemTotal > 0 && (
                      <div style={{ fontSize:10, color:C.textMuted, fontFamily:"monospace" }}>{fmtDollar(c.perDiemTotal)}</div>
                    )}
                  </td>
                  {/* Mileage */}
                  <td style={{ ...tdS, textAlign:"right" }}>{numIn(role.key,"mileageTrips",50,"0")}</td>
                  <td style={{ ...tdS, textAlign:"right" }}>{numIn(role.key,"milesPerTrip",60,"155")}</td>
                  {moneyCell(c.mileageCost)}
                  {/* Other travel */}
                  <td style={{ ...tdS, textAlign:"right" }}>{numIn(role.key,"travelTrips",50,"0")}</td>
                  <td style={{ ...tdS, textAlign:"right" }}>{numIn(role.key,"daysPerTrip",50,"0")}</td>
                  <td style={{ ...tdS, textAlign:"center" }}>{ynSel(role.key,"hotelRequired")}</td>
                  <td style={{ ...tdS, textAlign:"center" }}>{ynSel(role.key,"planeTicket")}</td>
                  <td style={{ ...tdS, textAlign:"center" }}>{ynSel(role.key,"carRental")}</td>
                  <td style={{ ...tdS, textAlign:"right", ...borderRight }}>{c.travelTotal > 0 ? <span style={{ fontFamily:"monospace", fontSize:11 }}>{fmtDollar(c.travelTotal)}</span> : <span style={{ color:C.textLight }}>—</span>}</td>
                  {/* Totals */}
                  <td style={{ ...tdS, textAlign:"right", fontFamily:"monospace", fontWeight:700, color: c.total > 0 ? C.accentLight : C.textLight }}>
                    {c.total > 0 ? fmtDollar(c.total) : "—"}
                  </td>
                  <td style={{ ...tdS, textAlign:"right", fontFamily:"monospace", fontSize:11, color:C.textMuted }}>
                    {c.loadedHourlyRate > 0 ? `$${c.loadedHourlyRate.toFixed(2)}/hr` : "—"}
                  </td>
                </tr>
              );
            })}

            {/* Subtotals row */}
            <tr style={{ backgroundColor:"#F0F4FF", fontWeight:700 }}>
              <td style={{ ...tdS, fontWeight:700, ...borderRight, position:"sticky", left:0, backgroundColor:"#F0F4FF", zIndex:1 }}>SUBTOTALS</td>
              <td colSpan={2} style={{ ...tdS }}></td>
              <td style={{ ...tdS, textAlign:"right", fontFamily:"monospace", ...borderRight }}>{fmtDollar(totals.salary)}</td>
              <td colSpan={2} style={{ ...tdS }}></td>
              <td style={{ ...tdS, textAlign:"right", fontFamily:"monospace", ...borderRight }}>{fmtDollar(totals.fleetFuel)}</td>
              <td colSpan={3} style={{ ...tdS, textAlign:"right", fontFamily:"monospace", ...borderRight }}>{fmtDollar(totals.equipment)}</td>
              <td style={{ ...tdS }}></td>
              <td style={{ ...tdS, textAlign:"right", fontFamily:"monospace", ...borderRight }}>{fmtDollar(totals.perDiem)}</td>
              <td colSpan={2} style={{ ...tdS }}></td>
              <td style={{ ...tdS, textAlign:"right", fontFamily:"monospace", ...borderRight }}>{fmtDollar(totals.mileage)}</td>
              <td colSpan={5} style={{ ...tdS }}></td>
              <td style={{ ...tdS, textAlign:"right", fontFamily:"monospace", ...borderRight }}>{fmtDollar(totals.travel)}</td>
              <td style={{ ...tdS, textAlign:"right", fontFamily:"monospace", fontSize:14, color:C.accentLight }}>{fmtDollar(totals.total)}</td>
              <td style={{ ...tdS }}></td>
            </tr>
          </tbody>
        </table>
      </div>

      {/* Footer — summary totals only */}
      <div style={{ padding:"10px 20px", borderTop:`2px solid ${C.border}`, backgroundColor:C.contentBg,
                    display:"flex", alignItems:"center", justifyContent:"flex-end", gap:40, flexShrink:0 }}>
        {[["Salaries", totals.salary], ["Per Diem", totals.perDiem], ["Travel & Mileage", totals.mileage + totals.travel], ["Total Supervision", totals.total]].map(([label, val]) => (
          <div key={label} style={{ textAlign:"right" }}>
            <div style={{ fontSize:11, color:C.textMuted, textTransform:"uppercase", letterSpacing:"0.06em" }}>{label}</div>
            <div style={{ fontSize: label === "Total Supervision" ? 16 : 13, fontWeight:700, fontFamily:"monospace", color:C.textPrimary }}>{fmtDollar(val)}</div>
          </div>
        ))}
      </div>


    </div>
  );
}

// ── BidSheet ──────────────────────────────────────────────────────────────────
function BidSheet({ bidData, onBidChange }) {
  function getRow(idx) { return bidData[idx] || {}; }
  function setField(idx, field, value) {
    onBidChange({ ...bidData, [idx]: { ...getRow(idx), [field]: value } });
  }

  const totalBid = BIDSHEET_TRADES.reduce((acc, _, idx) => {
    const amount = parseDollar(getRow(idx).amount ?? 0);
    return acc + amount;
  }, 0);

  const thStyle = {
    padding: "6px 10px", fontSize: 11, fontWeight: 700, color: C.textMuted,
    textTransform: "uppercase", letterSpacing: "0.06em",
    borderBottom: `2px solid ${C.border}`, textAlign: "left",
    backgroundColor: C.contentBg, position: "sticky", top: 0, zIndex: 2, whiteSpace: "nowrap",
  };
  const tdStyle = { padding: "4px 8px", fontSize: 12, borderBottom: `1px solid ${C.border}` };
  const inputStyle = { width: "100%", border: `1px solid ${C.border}`, borderRadius: 3, padding: "3px 6px", fontSize: 12, backgroundColor: C.white, boxSizing: "border-box" };
  const numStyle   = { ...inputStyle, fontFamily: "monospace", textAlign: "right" };

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100%", overflow: "hidden" }}>
      <div style={{ padding: "12px 20px 10px", borderBottom: `1px solid ${C.border}`, flexShrink: 0, display: "flex", alignItems: "baseline", gap: 12 }}>
        <span style={{ fontSize: 11, fontWeight: 700, color: C.textMuted, fontFamily: "monospace" }}>BIDSHEET</span>
        <span style={{ fontSize: 15, fontWeight: 700, color: C.textPrimary }}>Bid Sheet — Sub/Vendor Tracking</span>
        <span style={{ marginLeft: "auto", fontSize: 12, color: C.textMuted }}>
          TOTAL: <strong style={{ color: C.textPrimary, fontFamily: "monospace" }}>{fmtDollar(totalBid)}</strong>
        </span>
      </div>

      <div style={{ flex: 1, overflowY: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead>
            <tr>
              <th style={{ ...thStyle, width: 80  }}>CSI Code</th>
              <th style={{ ...thStyle            }}>Item</th>
              <th style={{ ...thStyle, width: 140, textAlign: "right" }}>Amount</th>
              <th style={{ ...thStyle, width: 200 }}>Sub / Vendor</th>
              <th style={{ ...thStyle, width: 60, textAlign: "center" }}>Type</th>
              <th style={{ ...thStyle, width: 80, textAlign: "center" }}>Subguard %</th>
            </tr>
          </thead>
          <tbody>
            {BIDSHEET_TRADES.map((trade, idx) => {
              const d      = getRow(idx);
              const amount = parseDollar(d.amount ?? 0);
              return (
                <tr key={idx} style={{ backgroundColor: idx % 2 === 0 ? C.white : C.tableStripe }}>
                  <td style={{ ...tdStyle, fontFamily: "monospace", fontSize: 11, color: C.textMuted }}>{trade.code}</td>
                  <td style={{ ...tdStyle }}>{trade.item}</td>
                  <td style={{ ...tdStyle, textAlign: "right" }}>
                    <input
                      style={numStyle}
                      type="number" min="0" step="0.01"
                      value={d.amount ?? ""}
                      onChange={e => setField(idx, "amount", e.target.value)}
                      placeholder="0.00"
                    />
                  </td>
                  <td style={{ ...tdStyle }}>
                    <input
                      style={inputStyle}
                      type="text"
                      value={d.vendor ?? ""}
                      onChange={e => setField(idx, "vendor", e.target.value)}
                      placeholder="Vendor name…"
                    />
                  </td>
                  <td style={{ ...tdStyle, textAlign: "center" }}>
                    <select
                      style={{ ...inputStyle, width: 60, textAlign: "center" }}
                      value={d.type ?? trade.type}
                      onChange={e => setField(idx, "type", e.target.value)}
                    >
                      {["S","M","No","NA"].map(t => <option key={t} value={t}>{t}</option>)}
                    </select>
                  </td>
                  <td style={{ ...tdStyle, textAlign: "center" }}>
                    <input
                      style={{ ...numStyle, width: 70 }}
                      type="number" min="0" max="100" step="0.1"
                      value={d.subguardPct ?? ""}
                      onChange={e => setField(idx, "subguardPct", e.target.value)}
                      placeholder="0"
                    />
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      <div style={{ padding: "10px 20px", borderTop: `2px solid ${C.border}`, backgroundColor: C.contentBg, display: "flex", justifyContent: "flex-end", flexShrink: 0 }}>
        <div style={{ textAlign: "right" }}>
          <div style={{ fontSize: 11, color: C.textMuted, textTransform: "uppercase", letterSpacing: "0.06em" }}>Total Bid</div>
          <div style={{ fontSize: 16, fontWeight: 700, fontFamily: "monospace", color: C.textPrimary }}>{fmtDollar(totalBid)}</div>
        </div>
      </div>
    </div>
  );
}

// ── RECAP Sheet ───────────────────────────────────────────────────────────────
const RECAP_TRADE_ROWS = [
  { code: "02 41", label: "DEMOLITION & EXISTING CONDITIONS" },
  { code: "03 00", label: "CONCRETE" },
  { code: "04 00", label: "MASONRY" },
  { code: "05 12", label: "STRUCTURAL STEEL / JOISTS / DECK" },
  { code: "05 50", label: "MISCELLANEOUS METALS" },
  { code: "06 10", label: "ROUGH CARPENTRY" },
  { code: "06 40", label: "MILLWORK & ARCHITECTURAL WOODWORK" },
  { code: "07 10", label: "DAMPPROOFING & WATERPROOFING" },
  { code: "07 24", label: "E.I.F.S." },
  { code: "07 40", label: "EXTERIOR WALL / METAL SIDING" },
  { code: "07 51", label: "ROOFING & SHEET METAL" },
  { code: "07 81", label: "SPRAYED FIREPROOFING" },
  { code: "08 11", label: "DOORS, FRAMES & HARDWARE" },
  { code: "08 41", label: "ENTRANCES, STOREFRONTS & CURTAIN WALL" },
  { code: "08 80", label: "GLASS, GLAZING & WINDOWS" },
  { code: "09 21", label: "DRYWALL & ACOUSTICS" },
  { code: "09 24", label: "PLASTER & STUCCO" },
  { code: "09 30", label: "TILE" },
  { code: "09 65", label: "RESILIENT FLOORING & CARPET" },
  { code: "09 90", label: "PAINTING & WALLCOVERING" },
  { code: "10 00", label: "SPECIALTIES" },
  { code: "11 00", label: "EQUIPMENT" },
  { code: "12 00", label: "FURNISHINGS" },
  { code: "13 00", label: "SPECIAL CONSTRUCTION" },
  { code: "14 20", label: "CONVEYING EQUIPMENT / ELEVATORS" },
  { code: "21 00", label: "FIRE SUPPRESSION" },
  { code: "22 00", label: "PLUMBING" },
  { code: "23 00", label: "HVAC" },
  { code: "26 00", label: "ELECTRICAL" },
  { code: "27 00", label: "COMMUNICATIONS" },
  { code: "28 00", label: "ELECTRONIC SAFETY & SECURITY" },
  { code: "31 20", label: "EARTHWORK & EXCAVATION" },
  { code: "31 63", label: "DRILLED PIERS" },
  { code: "32 12", label: "ASPHALT PAVING" },
  { code: "32 13", label: "SITE CONCRETE" },
  { code: "32 17", label: "PARKING LOT STRIPING" },
  { code: "32 31", label: "FENCES & GATES" },
  { code: "32 80", label: "LANDSCAPING & IRRIGATION" },
  { code: "33 10", label: "WATER UTILITIES" },
  { code: "33 30", label: "SANITARY SEWERAGE" },
  { code: "33 40", label: "STORM DRAINAGE" },
];

function RecapSheet({ project, recapData, onRecapChange, sheetTotals, firmDefaults, bidData, insSsRate }) {
  const sqft = parseDollar(project.sqft || 0);

  function getRow(code) { return recapData[code] || {}; }
  function setField(code, field, value) {
    onRecapChange({ ...recapData, [code]: { ...getRow(code), [field]: value } });
  }

  // Configurable fee/markup fields stored in recapData under special keys
  const feeRate        = parseDollar(recapData.__feeRate        ?? (firmDefaults.feeRate * 100));
  const bondRate       = parseDollar(recapData.__bondRate       ?? (firmDefaults.bondRate * 100));
  const glRate         = parseDollar(recapData.__glRate         ?? (firmDefaults.glRate * 100));
  const grtRate        = parseDollar(recapData.__grtRate        ?? (firmDefaults.grtRate * 100));
  const contingencyAmt = parseDollar(recapData.__contingencyAmt ?? 0);
  const insSsRateVal   = parseDollar(insSsRate ?? (firmDefaults.insSsRate * 100));

  function setConfig(key, value) {
    onRecapChange({ ...recapData, [key]: value });
  }

  // Aggregate BidSheet amounts by RECAP code.
  // BIDSHEET codes are 5-digit (e.g. "02050"); RECAP codes are 4-digit (e.g. "0205").
  // Match by stripping the last digit of the bidsheet code: bid "02050"[:4] = "0205".
  const bidByRecapCode = useMemo(() => {
    const map = {};
    BIDSHEET_TRADES.forEach((trade, idx) => {
      const d      = (bidData || {})[idx] || {};
      const amount = parseDollar(d.amount ?? 0);
      if (!amount) return;
      const type   = d.type ?? trade.type;
      // New CSI codes: "07 81 00" → first 5 chars = "07 81" matches recap key "07 81"
      // Also try first 2 chars (division) as fallback: "07 81" → "07"
      const key5 = trade.code.slice(0, 5); // "07 81"
      const key2 = trade.code.slice(0, 2); // "07"
      // Find the best matching recap row (longest prefix match)
      const recapKey = RECAP_TRADE_ROWS.find(r => r.code === key5)?.code
                    || RECAP_TRADE_ROWS.find(r => r.code.slice(0,2) === key2)?.code
                    || null;
      if (!recapKey) return;
      if (!map[recapKey]) map[recapKey] = { labor: 0, material: 0, sub: 0 };
      if (type === "M") map[recapKey].material += amount;
      else if (type === "S") map[recapKey].sub += amount;
    });
    return map;
  }, [bidData]);

  // Trade totals — bid amounts from BidSheet + manual overrides in recapData
  // Manual recapData entries override the bid-sourced value for that column.
  const tradeTotals = RECAP_TRADE_ROWS.reduce((acc, row) => {
    const d   = getRow(row.code);
    const bid = bidByRecapCode[row.code] || {};
    // If estimator has manually typed a value use it, otherwise use bid amount
    const labor    = parseDollar(d.labor    !== undefined && d.labor    !== "" ? d.labor    : 0);
    const material = parseDollar(d.material !== undefined && d.material !== "" ? d.material : (bid.material ?? 0));
    const sub      = parseDollar(d.sub      !== undefined && d.sub      !== "" ? d.sub      : (bid.sub      ?? 0));
    const total    = labor + material + sub;
    acc.labor    += labor;
    acc.material += material;
    acc.sub      += sub;
    acc.total    += total;
    return acc;
  }, { labor: 0, material: 0, sub: 0, total: 0 });

  const gcTotal    = (sheetTotals["2500"] || 0) + (sheetTotals["2600"] || 0) +
                     (sheetTotals["2700"] || 0) + (sheetTotals["2900"] || 0);
  const subtotal   = tradeTotals.total + gcTotal;
  const insSsAmt   = tradeTotals.labor * (insSsRateVal / 100);
  const subtotal2  = subtotal + insSsAmt;
  const feeAmt     = subtotal2 * (feeRate / 100);
  const bondAmt    = (subtotal2 + feeAmt) * (bondRate / 100);
  const glAmt      = (subtotal2 + feeAmt) * (glRate / 100);
  const grtAmt     = (subtotal2 + feeAmt) * (grtRate / 100);
  const projectTotal = subtotal2 + feeAmt + bondAmt + glAmt + grtAmt + contingencyAmt;

  const thStyle = {
    padding: "6px 10px", fontSize: 10, fontWeight: 700, color: C.textMuted,
    textTransform: "uppercase", letterSpacing: "0.06em",
    borderBottom: `2px solid ${C.border}`, textAlign: "right",
    backgroundColor: C.contentBg, position: "sticky", top: 0, zIndex: 2, whiteSpace: "nowrap",
  };
  const tdStyle  = { padding: "5px 10px", fontSize: 12, borderBottom: `1px solid ${C.border}`, textAlign: "right", fontFamily: "monospace" };
  const numInput = (code, field) => {
    const bid = bidByRecapCode[code] || {};
    const bidVal = field === "material" ? (bid.material ?? 0) : field === "sub" ? (bid.sub ?? 0) : 0;
    const hasOverride = getRow(code)[field] !== undefined && getRow(code)[field] !== "";
    return (
      <input
        style={{
          width: 120, border: `1px solid ${hasOverride ? C.border : (bidVal > 0 ? "#BFDBFE" : C.border)}`,
          borderRadius: 3, padding: "3px 8px", fontSize: 12, fontFamily: "monospace",
          backgroundColor: hasOverride ? C.white : (bidVal > 0 ? "#EFF6FF" : C.white),
          textAlign: "right", boxSizing: "border-box", color: C.textPrimary,
        }}
        type="number" min="0" step="0.01"
        value={getRow(code)[field] ?? ""}
        onChange={e => setField(code, field, e.target.value)}
        placeholder={bidVal > 0 ? bidVal.toFixed(2) : "0.00"}
      />
    );
  };


  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100%", overflow: "hidden" }}>
      {/* Single consolidated header */}
      <div style={{ padding: "12px 20px", borderBottom: `1px solid ${C.border}`, flexShrink: 0, backgroundColor: C.white }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10 }}>
          <span style={{ fontSize: 15, fontWeight: 700, color: C.textPrimary }}>Project Recap</span>
          <div style={{ textAlign: "right" }}>
            <div style={{ fontSize: 11, color: C.textMuted, textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 2 }}>Project Total</div>
            <div style={{ fontSize: 18, fontWeight: 700, fontFamily: "monospace", color: C.textPrimary }}>
              {fmtDollar(projectTotal)}
              {sqft > 0 && <span style={{ fontSize: 12, fontWeight: 400, color: C.textMuted, marginLeft: 10 }}>{fmtDollar(projectTotal / sqft)}/SF</span>}
            </div>
          </div>
        </div>
        <div style={{ display: "flex", gap: 28, flexWrap: "wrap" }}>
          {[
            ["Job No.",    project.jobNo],
            ["Job Name",   project.jobName],
            ["Location",   project.location || "—"],
            ["Area",       project.sqft ? Number(project.sqft).toLocaleString() + " SF" : "—"],
            ["Estimator",  project.estimator],
            ["Duration",   project.duration + " months"],
          ].map(([label, val]) => (
            <div key={label}>
              <div style={{ fontSize: 10, fontWeight: 600, color: C.textMuted, textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 2 }}>{label}</div>
              <div style={{ fontSize: 12, color: C.textPrimary, fontWeight: label === "Job No." ? 600 : 400 }}>{val}</div>
            </div>
          ))}
        </div>
      </div>

      <div style={{ flex: 1, overflowY: "auto" }}>

        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead>
            <tr>
              <th style={{ ...thStyle, textAlign: "left", width: 70 }}>Code</th>
              <th style={{ ...thStyle, textAlign: "left" }}>Description</th>
              <th style={{ ...thStyle, width: 130 }}>Labor</th>
              <th style={{ ...thStyle, width: 120 }}>Material</th>
              <th style={{ ...thStyle, width: 130 }}>Sub / Equip</th>
              <th style={{ ...thStyle, width: 130 }}>Total</th>
              <th style={{ ...thStyle, width: 80 }}>$/SQFT</th>
            </tr>
          </thead>
          <tbody>
            {/* Trade rows */}
            {RECAP_TRADE_ROWS.map((row, i) => {
              const d        = getRow(row.code);
              const bid      = bidByRecapCode[row.code] || {};
              const labor    = parseDollar(d.labor    !== undefined && d.labor    !== "" ? d.labor    : 0);
              const material = parseDollar(d.material !== undefined && d.material !== "" ? d.material : (bid.material ?? 0));
              const sub      = parseDollar(d.sub      !== undefined && d.sub      !== "" ? d.sub      : (bid.sub      ?? 0));
              const total    = labor + material + sub;
              return (
                <tr key={row.code} style={{ backgroundColor: i % 2 === 0 ? C.white : C.tableStripe }}>
                  <td style={{ ...tdStyle, textAlign: "left", fontFamily: "monospace", fontSize: 11, color: C.textMuted }}>{row.code}</td>
                  <td style={{ ...tdStyle, textAlign: "left", fontFamily: "Inter, sans-serif", color: C.textPrimary }}>{row.label}</td>
                  <td style={{ ...tdStyle }}>{numInput(row.code, "labor")}</td>
                  <td style={{ ...tdStyle }}>{numInput(row.code, "material")}</td>
                  <td style={{ ...tdStyle }}>{numInput(row.code, "sub")}</td>
                  <td style={{ ...tdStyle, fontWeight: total > 0 ? 700 : 400, color: total > 0 ? C.textPrimary : C.textLight }}>{total > 0 ? fmtDollar(total) : "—"}</td>
                  <td style={{ ...tdStyle, color: total > 0 ? C.textMuted : C.textLight }}>{total > 0 && sqft > 0 ? fmtDollar(total / sqft) : "—"}</td>
                </tr>
              );
            })}

            {/* GC Sheets sub-totals (read from live sheet calcs) */}
            {[
              { code: "2500", label: "GENERAL CONDITIONS", val: sheetTotals["2500"] || 0 },
              { code: "2600", label: "EQUIPMENT",           val: sheetTotals["2600"] || 0 },
              { code: "2700", label: "INSURANCE, PERMITS & BONDING", val: sheetTotals["2700"] || 0 },
              { code: "2900", label: "ALLOWANCES / TESTING & INSPECTION", val: sheetTotals["2900"] || 0 },
            ].map((row, i) => (
              <tr key={row.code} style={{ backgroundColor: (RECAP_TRADE_ROWS.length + i) % 2 === 0 ? C.white : C.tableStripe }}>
                <td style={{ ...tdStyle, textAlign: "left", fontFamily: "monospace", fontSize: 11, color: C.accentLight }}>{row.code}</td>
                <td style={{ ...tdStyle, textAlign: "left", fontFamily: "Inter, sans-serif", color: C.textPrimary, fontStyle: "italic" }}>{row.label}</td>
                <td style={{ ...tdStyle, color: C.textLight }}>—</td>
                <td style={{ ...tdStyle, color: C.textLight }}>—</td>
                <td style={{ ...tdStyle, fontWeight: row.val > 0 ? 600 : 400, color: row.val > 0 ? C.textPrimary : C.textLight }}>{row.val > 0 ? fmtDollar(row.val) : "—"}</td>
                <td style={{ ...tdStyle, fontWeight: row.val > 0 ? 700 : 400, color: row.val > 0 ? C.textPrimary : C.textLight }}>{row.val > 0 ? fmtDollar(row.val) : "—"}</td>
                <td style={{ ...tdStyle, color: C.textLight }}>{row.val > 0 && sqft > 0 ? fmtDollar(row.val / sqft) : "—"}</td>
              </tr>
            ))}

            {/* Subtotal */}
            <tr style={{ backgroundColor: "#F0F4FF" }}>
              <td colSpan={2} style={{ ...tdStyle, textAlign: "left", fontWeight: 700, fontFamily: "Inter, sans-serif", fontSize: 12, borderTop: `2px solid ${C.border}` }}>SUBTOTAL</td>
              <td style={{ ...tdStyle, fontWeight: 700, borderTop: `2px solid ${C.border}` }}>{fmtDollar(tradeTotals.labor)}</td>
              <td style={{ ...tdStyle, fontWeight: 700, borderTop: `2px solid ${C.border}` }}>{fmtDollar(tradeTotals.material)}</td>
              <td style={{ ...tdStyle, fontWeight: 700, borderTop: `2px solid ${C.border}` }}>{fmtDollar(tradeTotals.sub + gcTotal)}</td>
              <td style={{ ...tdStyle, fontWeight: 700, borderTop: `2px solid ${C.border}` }}>{fmtDollar(subtotal)}</td>
              <td style={{ ...tdStyle, fontWeight: 700, borderTop: `2px solid ${C.border}` }}>{sqft > 0 ? fmtDollar(subtotal/sqft) : "—"}</td>
            </tr>

            {/* Fee / Bond / GL / GRT / Contingency — inline editable % in description cell */}
            {[
              { code: "5000-002", label: "GC FEE",                  rateKey: "__feeRate",  rate: feeRate,  val: feeAmt,         isAmt: false },
              { code: "2700-280", label: "GC BOND",                  rateKey: "__bondRate", rate: bondRate, val: bondAmt,        isAmt: false },
              { code: "2700-350", label: "GENERAL LIABILITY INS",    rateKey: "__glRate",   rate: glRate,   val: glAmt,          isAmt: false },
              { code: "2800-100", label: "GROSS RECEIPTS TAX (GRT)", rateKey: "__grtRate",  rate: grtRate,  val: grtAmt,         isAmt: false },
              { code: "5000-004", label: "CONTINGENCY",              rateKey: "__contingencyAmt", rate: contingencyAmt, val: contingencyAmt, isAmt: true  },
            ].map(({ code, label, rateKey, rate, val, isAmt }) => (
              <tr key={code} style={{ backgroundColor: C.white }}>
                <td style={{ ...tdStyle, textAlign: "left", fontFamily: "monospace", fontSize: 11, color: C.textMuted }}>{code}</td>
                <td style={{ ...tdStyle, textAlign: "left" }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                    <span style={{ fontFamily: "Inter, sans-serif", color: C.textPrimary, fontSize: 12 }}>{label}</span>
                    <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
                      <input
                        style={{ width: 80, border: `1px solid ${C.border}`, borderRadius: 3, padding: "2px 6px", fontSize: 11, fontFamily: "monospace", backgroundColor: "#F8FAFF", textAlign: "right", boxSizing: "border-box" }}
                        type="number" min="0" step="0.01"
                        value={recapData[rateKey] ?? ""}
                        onChange={e => setConfig(rateKey, e.target.value)}
                        placeholder={isAmt ? "0.00" : rate.toFixed(2)}
                      />
                      <span style={{ fontSize: 11, color: C.textMuted }}>{isAmt ? "$" : "%"}</span>
                    </div>
                  </div>
                </td>
                <td colSpan={3} style={{ ...tdStyle }}></td>
                <td style={{ ...tdStyle, fontWeight: 600, color: val > 0 ? C.textPrimary : C.textLight }}>{val > 0 ? fmtDollar(val) : "—"}</td>
                <td style={{ ...tdStyle, color: C.textMuted }}>{val > 0 && sqft > 0 ? fmtDollar(val/sqft) : "—"}</td>
              </tr>
            ))}

            {/* Project Total */}
            <tr style={{ backgroundColor: C.accent }}>
              <td colSpan={2} style={{ ...tdStyle, textAlign: "left", fontWeight: 700, fontFamily: "Inter, sans-serif", fontSize: 13, color: C.white, borderTop: `2px solid ${C.border}` }}>PROJECT TOTAL</td>
              <td colSpan={3} style={{ ...tdStyle, borderTop: `2px solid ${C.border}` }}></td>
              <td style={{ ...tdStyle, fontWeight: 800, fontSize: 14, color: C.white, borderTop: `2px solid ${C.border}` }}>{fmtDollar(projectTotal)}</td>
              <td style={{ ...tdStyle, fontWeight: 700, color: "#94A3B8", borderTop: `2px solid ${C.border}` }}>{sqft > 0 ? fmtDollar(projectTotal/sqft) : "—"}</td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ── Firm Settings Modal ───────────────────────────────────────────────────────

// ── Schedule of Values (SOV) ──────────────────────────────────────────────────
function SOVSheet({ project, recapData, bidData, sheetTotals, firmDefaults, insSsRate }) {
  const sqft       = parseDollar(project.sqft || 0);
  const today      = new Date().toLocaleDateString("en-US", { month: "long", day: "numeric", year: "numeric" });

  // ── Replicate RECAP calculations ──────────────────────────────────────────
  function getRow(code) { return recapData[code] || {}; }

  const bidByCode = useMemo(() => {
    const map = {};
    BIDSHEET_TRADES.forEach((trade, idx) => {
      const d      = (bidData || {})[idx] || {};
      const amount = parseDollar(d.amount ?? 0);
      if (!amount) return;
      const type   = d.type ?? trade.type;
      const key5   = trade.code.slice(0, 5);
      const key2   = trade.code.slice(0, 2);
      const recapKey = RECAP_TRADE_ROWS.find(r => r.code === key5)?.code
                    || RECAP_TRADE_ROWS.find(r => r.code.slice(0,2) === key2)?.code
                    || null;
      if (!recapKey) return;
      if (!map[recapKey]) map[recapKey] = { labor: 0, material: 0, sub: 0 };
      if (type === "M") map[recapKey].material += amount;
      else if (type === "S") map[recapKey].sub += amount;
    });
    return map;
  }, [bidData]);

  const feeRate        = parseDollar(recapData.__feeRate        ?? (firmDefaults.feeRate * 100));
  const bondRate       = parseDollar(recapData.__bondRate       ?? (firmDefaults.bondRate * 100));
  const glRate         = parseDollar(recapData.__glRate         ?? (firmDefaults.glRate * 100));
  const grtRate        = parseDollar(recapData.__grtRate        ?? (firmDefaults.grtRate * 100));
  const contingencyAmt = parseDollar(recapData.__contingencyAmt ?? 0);
  const insSsRateVal   = parseDollar(insSsRate ?? (firmDefaults.insSsRate * 100));

  const tradeLines = RECAP_TRADE_ROWS.map(row => {
    const d   = getRow(row.code);
    const bid = bidByCode[row.code] || {};
    const labor    = parseDollar(d.labor    !== undefined && d.labor    !== "" ? d.labor    : 0);
    const material = parseDollar(d.material !== undefined && d.material !== "" ? d.material : (bid.material ?? 0));
    const sub      = parseDollar(d.sub      !== undefined && d.sub      !== "" ? d.sub      : (bid.sub      ?? 0));
    const total    = labor + material + sub;
    return { ...row, labor, material, sub, total };
  }).filter(r => r.total > 0);

  const gcTotal    = (sheetTotals["2500"] || 0) + (sheetTotals["2600"] || 0) +
                     (sheetTotals["2700"] || 0) + (sheetTotals["2900"] || 0);
  const tradeTotal = tradeLines.reduce((a, r) => a + r.total, 0);
  const subtotal   = tradeTotal + gcTotal;
  const insSsAmt   = tradeLines.reduce((a, r) => a + r.labor, 0) * (insSsRateVal / 100);
  const subtotal2  = subtotal + insSsAmt;
  const feeAmt     = subtotal2 * (feeRate  / 100);
  const bondAmt    = (subtotal2 + feeAmt) * (bondRate / 100);
  const glAmt      = (subtotal2 + feeAmt) * (glRate   / 100);
  const grtAmt     = (subtotal2 + feeAmt) * (grtRate  / 100);
  const projectTotal = subtotal2 + feeAmt + bondAmt + glAmt + grtAmt + contingencyAmt;

  // ── Print styles injected once ────────────────────────────────────────────
  const printStyle = `
    @media print {
      body > * { display: none !important; }
      #sov-print-root { display: block !important; }
      #sov-print-root { position: fixed; top: 0; left: 0; width: 100%; }
    }
  `;

  const accent  = "#0F1923";   // navy — Meridian brand
  const gold    = "#B8933A";   // accent gold for dividers
  const light   = "#F7F8FA";
  const border  = "#E2E8F0";

  const thStyle = {
    padding: "7px 14px", fontSize: 10, fontWeight: 700,
    textTransform: "uppercase", letterSpacing: "0.08em",
    color: C.white, backgroundColor: accent,
    textAlign: "right", whiteSpace: "nowrap",
  };
  const tdRight = { padding: "7px 14px", fontSize: 12, fontFamily: "monospace", textAlign: "right", borderBottom: `1px solid ${border}` };
  const tdLeft  = { padding: "7px 14px", fontSize: 12, textAlign: "left", borderBottom: `1px solid ${border}` };

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100%", overflow: "hidden" }}>
      <style>{printStyle}</style>

      {/* Toolbar */}
      <div style={{ padding: "8px 20px", borderBottom: `1px solid ${C.border}`, flexShrink: 0,
                    display: "flex", alignItems: "center", justifyContent: "space-between", backgroundColor: C.white }}>
        <span style={{ fontSize: 12, color: C.textMuted }}>
          Schedule of Values — {project.jobName}
          {projectTotal > 0
            ? <strong style={{ color: C.textPrimary, marginLeft: 10, fontFamily: "monospace" }}>{fmtDollar(projectTotal)}</strong>
            : <span style={{ marginLeft: 10, color: C.textLight }}>No costs entered yet</span>}
        </span>
        <button style={{ ...s.btn(), fontSize: 12, padding: "3px 12px" }} onClick={() => window.print()}>
          ⎙ Print / Export PDF
        </button>
      </div>

      {/* Scrollable document */}
      <div id="sov-print-root" style={{ flex: 1, overflowY: "auto", backgroundColor: "#E8EAED", padding: "32px 40px" }}>
        <div style={{
          maxWidth: 900, margin: "0 auto",
          backgroundColor: C.white,
          boxShadow: "0 4px 24px rgba(0,0,0,0.12)",
          fontFamily: "Inter, sans-serif",
        }}>

          {/* ── COVER HEADER ─────────────────────────────────────────────── */}
          <div style={{ backgroundColor: accent, padding: "48px 56px 36px" }}>
            {/* Logo / firm name */}
            <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", marginBottom: 40 }}>
              <div>
                <div style={{ fontSize: 11, fontWeight: 700, color: gold, textTransform: "uppercase", letterSpacing: "0.15em", marginBottom: 6 }}>
                  General Contractor
                </div>
                <div style={{ fontSize: 26, fontWeight: 800, color: C.white, letterSpacing: "-0.02em", lineHeight: 1.1 }}>
                  Meridian Construction Group
                </div>
              </div>
              {/* Placeholder logo box */}
              <div style={{ width: 72, height: 72, border: `2px solid rgba(255,255,255,0.2)`, borderRadius: 4,
                            display: "flex", alignItems: "center", justifyContent: "center" }}>
                <span style={{ fontSize: 22, fontWeight: 900, color: gold }}>MCG</span>
              </div>
            </div>

            {/* Document title */}
            <div style={{ borderTop: `1px solid ${gold}`, paddingTop: 24 }}>
              <div style={{ fontSize: 11, fontWeight: 600, color: gold, textTransform: "uppercase", letterSpacing: "0.14em", marginBottom: 8 }}>
                Proposal — Schedule of Values
              </div>
              <div style={{ fontSize: 32, fontWeight: 800, color: C.white, lineHeight: 1.15, letterSpacing: "-0.01em" }}>
                {project.jobName}
              </div>
              {project.location && (
                <div style={{ fontSize: 14, color: "rgba(255,255,255,0.65)", marginTop: 6 }}>{project.location}</div>
              )}
            </div>
          </div>

          {/* ── PROJECT METADATA ─────────────────────────────────────────── */}
          <div style={{ backgroundColor: light, borderBottom: `3px solid ${gold}`, padding: "20px 56px" }}>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 24 }}>
              {[
                ["Project No.",    project.jobNo],
                ["Project Type",   project.projectType],
                ["Building Area",  sqft > 0 ? Number(sqft).toLocaleString() + " SF" : "—"],
                ["Duration",       project.duration + " Months"],
                ["Estimator",      project.estimator],
                ["Date Prepared",  today],
                ["Proposal Total", fmtDollar(projectTotal)],
                ["Cost per SF",    sqft > 0 ? fmtDollar(projectTotal / sqft) + " / SF" : "—"],
              ].map(([label, val]) => (
                <div key={label}>
                  <div style={{ fontSize: 9, fontWeight: 700, color: "#94A3B8", textTransform: "uppercase", letterSpacing: "0.1em", marginBottom: 3 }}>{label}</div>
                  <div style={{ fontSize: 13, fontWeight: label === "Proposal Total" ? 700 : 500, color: accent,
                                fontFamily: ["Proposal Total","Cost per SF","Building Area"].includes(label) ? "monospace" : "inherit" }}>
                    {val}
                  </div>
                </div>
              ))}
            </div>
          </div>

          {/* ── SOV TABLE ────────────────────────────────────────────────── */}
          <div style={{ padding: "32px 0 0" }}>
            <div style={{ padding: "0 56px 12px", display: "flex", alignItems: "baseline", gap: 12 }}>
              <div style={{ fontSize: 14, fontWeight: 700, color: accent, textTransform: "uppercase", letterSpacing: "0.08em" }}>
                Schedule of Values
              </div>
              <div style={{ flex: 1, height: 1, backgroundColor: gold, marginLeft: 8 }} />
            </div>

            <table style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead>
                <tr>
                  <th style={{ ...thStyle, textAlign: "left", width: 80, paddingLeft: 56 }}>CSI Code</th>
                  <th style={{ ...thStyle, textAlign: "left" }}>Description</th>
                  <th style={{ ...thStyle }}>Total</th>
                  <th style={{ ...thStyle, paddingRight: 56 }}>$/SF</th>
                </tr>
              </thead>
              <tbody>
                {/* Trade lines — only non-zero */}
                {tradeLines.length === 0 ? (
                  <tr>
                    <td colSpan={4} style={{ ...tdLeft, paddingLeft: 56, color: C.textMuted, fontStyle: "italic", textAlign: "center", padding: "24px" }}>
                      No trade costs entered. Complete the Recap tab to populate this schedule.
                    </td>
                  </tr>
                ) : tradeLines.map((row, i) => (
                  <tr key={row.code} style={{ backgroundColor: i % 2 === 0 ? C.white : light }}>
                    <td style={{ ...tdLeft, paddingLeft: 56, fontFamily: "monospace", fontSize: 11, color: "#64748B" }}>{row.code}</td>
                    <td style={{ ...tdLeft, fontWeight: 500, color: accent }}>{row.label}</td>
                    <td style={{ ...tdRight, fontWeight: 600, color: accent }}>{fmtDollar(row.total)}</td>
                    <td style={{ ...tdRight, color: "#64748B", paddingRight: 56 }}>{sqft > 0 ? fmtDollar(row.total / sqft) : "—"}</td>
                  </tr>
                ))}

                {/* General Conditions — one rolled-up line */}
                {gcTotal > 0 && (
                  <tr style={{ backgroundColor: tradeLines.length % 2 === 0 ? C.white : light }}>
                    <td style={{ ...tdLeft, paddingLeft: 56, fontFamily: "monospace", fontSize: 11, color: "#64748B" }}>01 00 00</td>
                    <td style={{ ...tdLeft, fontWeight: 500, color: accent }}>GENERAL CONDITIONS</td>
                    <td style={{ ...tdRight, fontWeight: 600, color: accent }}>{fmtDollar(gcTotal)}</td>
                    <td style={{ ...tdRight, color: "#64748B", paddingRight: 56 }}>{sqft > 0 ? fmtDollar(gcTotal / sqft) : "—"}</td>
                  </tr>
                )}

                {/* Subtotal */}
                <tr style={{ backgroundColor: "#F0F4FF" }}>
                  <td colSpan={2} style={{ ...tdLeft, paddingLeft: 56, fontWeight: 700, fontSize: 13, color: accent, borderTop: `2px solid ${border}` }}>SUBTOTAL</td>
                  <td style={{ ...tdRight, fontWeight: 700, fontSize: 13, color: accent, borderTop: `2px solid ${border}` }}>{fmtDollar(subtotal2)}</td>
                  <td style={{ ...tdRight, fontWeight: 600, color: "#64748B", paddingRight: 56, borderTop: `2px solid ${border}` }}>{sqft > 0 ? fmtDollar(subtotal2/sqft) : "—"}</td>
                </tr>

                {/* Fee / Bond / GL / GRT / Contingency — only non-zero */}
                {[
                  ["", `GC Fee (${feeRate.toFixed(2)}%)`,                  feeAmt],
                  ["", `Performance & Payment Bond (${bondRate.toFixed(2)}%)`, bondAmt],
                  ["", `General Liability Insurance (${glRate.toFixed(2)}%)`,  glAmt],
                  ["", `Gross Receipts Tax (${grtRate.toFixed(2)}%)`,       grtAmt],
                  ["", "Contingency",                                         contingencyAmt],
                ].filter(([,, val]) => val > 0).map(([code, label, val], i) => (
                  <tr key={label} style={{ backgroundColor: C.white }}>
                    <td style={{ ...tdLeft, paddingLeft: 56, fontFamily: "monospace", fontSize: 11, color: "#94A3B8" }}>{code}</td>
                    <td style={{ ...tdLeft, color: "#64748B", fontStyle: "italic" }}>{label}</td>
                    <td style={{ ...tdRight, color: accent }}>{fmtDollar(val)}</td>
                    <td style={{ ...tdRight, color: "#94A3B8", paddingRight: 56 }}>{sqft > 0 ? fmtDollar(val/sqft) : "—"}</td>
                  </tr>
                ))}

                {/* Project Total */}
                <tr style={{ backgroundColor: accent }}>
                  <td colSpan={2} style={{ ...tdLeft, paddingLeft: 56, fontWeight: 800, fontSize: 15, color: C.white, borderTop: `3px solid ${gold}`, padding: "14px 14px 14px 56px" }}>
                    TOTAL PROPOSED CONTRACT SUM
                  </td>
                  <td style={{ ...tdRight, fontWeight: 800, fontSize: 15, color: C.white, borderTop: `3px solid ${gold}`, padding: "14px" }}>
                    {fmtDollar(projectTotal)}
                  </td>
                  <td style={{ ...tdRight, fontWeight: 700, color: gold, borderTop: `3px solid ${gold}`, paddingRight: 56, padding: "14px 56px 14px 14px" }}>
                    {sqft > 0 ? fmtDollar(projectTotal/sqft) : "—"}
                  </td>
                </tr>
              </tbody>
            </table>
          </div>

          {/* ── FOOTER / SIGNATURE BLOCK ─────────────────────────────────── */}
          <div style={{ padding: "40px 56px 48px", borderTop: `1px solid ${border}`, marginTop: 32 }}>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 48 }}>
              {/* Clarifications */}
              <div>
                <div style={{ fontSize: 10, fontWeight: 700, color: gold, textTransform: "uppercase", letterSpacing: "0.12em", marginBottom: 10 }}>
                  Clarifications & Assumptions
                </div>
                {[
                  "This proposal is based on documents as listed and described above.",
                  "All pricing is subject to review and confirmation of final contract documents.",
                  "Pricing is valid for 30 days from the date of this proposal.",
                  "This proposal does not include Owner-furnished or Owner-installed items.",
                  "Applicable taxes, permits, and fees are included as noted.",
                ].map((line, i) => (
                  <div key={i} style={{ fontSize: 11, color: "#475569", marginBottom: 5, paddingLeft: 12, position: "relative" }}>
                    <span style={{ position: "absolute", left: 0, color: gold }}>·</span>
                    {line}
                  </div>
                ))}
              </div>

              {/* Signature */}
              <div>
                <div style={{ fontSize: 10, fontWeight: 700, color: gold, textTransform: "uppercase", letterSpacing: "0.12em", marginBottom: 10 }}>
                  Submitted By
                </div>
                <div style={{ fontSize: 13, fontWeight: 600, color: accent, marginBottom: 2 }}>Meridian Construction Group</div>
                <div style={{ fontSize: 12, color: "#64748B", marginBottom: 24 }}>{project.estimator}</div>
                <div style={{ borderBottom: `1px solid ${accent}`, width: 220, marginBottom: 4 }} />
                <div style={{ fontSize: 10, color: "#94A3B8" }}>Authorized Signature &amp; Date</div>
                <div style={{ borderBottom: `1px solid ${accent}`, width: 220, marginTop: 24, marginBottom: 4 }} />
                <div style={{ fontSize: 10, color: "#94A3B8" }}>Accepted By &amp; Date</div>
              </div>
            </div>

            {/* Firm footer */}
            <div style={{ marginTop: 36, paddingTop: 16, borderTop: `1px solid ${border}`,
                          display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <div style={{ fontSize: 10, color: "#94A3B8" }}>
                Meridian Construction Group &nbsp;·&nbsp; Confidential — For Proposal Purposes Only
              </div>
              <div style={{ fontSize: 10, color: "#94A3B8" }}>{today}</div>
            </div>
          </div>

        </div>
      </div>
    </div>
  );
}

const SUPERVISION_ASSUMPTION_DEFAULTS = {
  hoursPerWeek: 40, weeksPerMonth: 4.33, mileageRate: 0.725,
  hotelRate: 160, planeRate: 600, carRentalRate: 150,
  fuelPerGal: 4.15, mpg: 12, defaultDailyMiles: 50, foodAllowance: 50,
};

function SettingsModal({ firmRates, assumptions, insSsRate, onSave, onClose }) {
  const [rates,      setRates]      = useState({ ...firmRates });
  const [draftAssump,setDraftAssump]= useState({ ...assumptions });
  const [draftInsSs, setDraftInsSs] = useState(insSsRate ?? "28.00");

  return (
    <div style={s.overlay} onClick={e => e.target === e.currentTarget && onClose()}>
      <div style={{ backgroundColor: C.white, borderRadius: 4, width: 720, maxHeight: "88vh", display: "flex", flexDirection: "column", boxShadow: "0 16px 48px rgba(0,0,0,0.25)" }}>
        <div style={{ ...s.modalHeader, borderBottom: `1px solid ${C.border}` }}>
          <div>
            <div style={s.modalTitle}>Assumptions</div>
            <div style={s.modalSubtitle}>Supervision rates, assumptions, and labor burden</div>
          </div>
          <button style={s.closeBtn} onClick={onClose}>×</button>
        </div>
        <div style={{ ...s.modalBody, overflowY: "auto", flex: 1 }}>

          <div style={s.sectionDivider}>Labor Burden</div>
          <div style={s.grid3}>
            <div style={s.field}>
              <label style={s.label}>Ins & SS Burden Rate (%)</label>
              <input style={{ ...s.input, fontFamily:"monospace" }} type="number" min="0" step="0.01"
                value={draftInsSs}
                placeholder="28.00"
                onChange={e => setDraftInsSs(e.target.value)} />
            </div>
          </div>

          <div style={s.sectionDivider}>Time Conversion</div>
          <div style={s.grid3}>
            {[
              ["Hours / Week",  "hoursPerWeek",  "40"],
              ["Weeks / Month", "weeksPerMonth", "4.33"],
            ].map(([label, key, ph]) => (
              <div key={key} style={s.field}>
                <label style={s.label}>{label}</label>
                <input style={{ ...s.input, fontFamily:"monospace" }} type="number" min="0" step="0.01"
                  value={draftAssump[key] ?? ""} placeholder={ph}
                  onChange={e => setDraftAssump(p => ({ ...p, [key]: parseFloat(e.target.value)||0 }))} />
              </div>
            ))}
          </div>

          <div style={s.sectionDivider}>Vehicle & Fuel</div>
          <div style={s.grid3}>
            {[
              ["Fuel $/Gal",           "fuelPerGal",        "4.15"],
              ["MPG (Fleet Vehicles)", "mpg",               "12"],
              ["Default Daily Miles",  "defaultDailyMiles", "50"],
            ].map(([label, key, ph]) => (
              <div key={key} style={s.field}>
                <label style={s.label}>{label}</label>
                <input style={{ ...s.input, fontFamily:"monospace" }} type="number" min="0" step="0.01"
                  value={draftAssump[key] ?? ""} placeholder={ph}
                  onChange={e => setDraftAssump(p => ({ ...p, [key]: parseFloat(e.target.value)||0 }))} />
              </div>
            ))}
          </div>

          <div style={s.sectionDivider}>Mileage & Travel</div>
          <div style={s.grid3}>
            {[
              ["Mileage Rate ($/mi)",    "mileageRate",   "0.725"],
              ["Avg Hotel ($/night)",    "hotelRate",     "160"],
              ["Avg Plane ($/roundtrip)","planeRate",     "600"],
              ["Car Rental ($/day)",     "carRentalRate", "150"],
              ["Food Allowance ($/day)", "foodAllowance", "50"],
            ].map(([label, key, ph]) => (
              <div key={key} style={s.field}>
                <label style={s.label}>{label}</label>
                <input style={{ ...s.input, fontFamily:"monospace" }} type="number" min="0" step="0.01"
                  value={draftAssump[key] ?? ""} placeholder={ph}
                  onChange={e => setDraftAssump(p => ({ ...p, [key]: parseFloat(e.target.value)||0 }))} />
              </div>
            ))}
          </div>

          <div style={s.sectionDivider}>Role Hourly Rates ($/hr)</div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
            {SUPERVISION_ROLES.map(role => (
              <div key={role.key} style={s.field}>
                <label style={s.label}>{role.label}</label>
                <input style={{ ...s.input, fontFamily: "monospace" }} type="number" min="0" step="0.01"
                  value={rates[role.key] ?? role.defaultRate}
                  onChange={e => setRates(p => ({ ...p, [role.key]: parseFloat(e.target.value) || 0 }))} />
              </div>
            ))}
          </div>
        </div>
        <div style={{ ...s.modalFooter, borderTop: `1px solid ${C.border}`, justifyContent: "space-between" }}>
          <button style={s.btn()} onClick={() => setDraftAssump({ ...SUPERVISION_ASSUMPTION_DEFAULTS })}>Reset Assumptions</button>
          <div style={{ display: "flex", gap: 8 }}>
            <button style={s.btn()} onClick={onClose}>Cancel</button>
            <button style={s.btn("primary")} onClick={() => { onSave(rates, draftAssump, draftInsSs); onClose(); }}>Save</button>
          </div>
        </div>
      </div>
    </div>
  );
}

// ── New Project Modal ─────────────────────────────────────────────────────────
const INDUSTRIES = ["Retail","Restaurant","Education","Healthcare","Municipal","C-Store","Office","Industrial","Multifamily","Hospitality","Mixed-Use","Other"];
const BID_TYPES  = ["Hard Bid","Negotiated","Design-Build","CM at Risk"];

function NewProjectModal({ onClose, onCreate }) {
  const [form, setForm] = useState({
    jobNo: `MCG-${new Date().getFullYear()}-${String(Math.floor(Math.random()*900)+100)}`,
    jobName: "", location: "", sqft: "", estimator: MOCK_USER.name,
    duration: "", projectType: "Commercial",
    client: "", industry: "Retail", bidType: "Hard Bid",
    projectedFee: "", dueDate: "", contactName: "", contactEmail: "",
  });
  const [error, setError] = useState("");
  const f = k => e => setForm(p => ({ ...p, [k]: e.target.value }));

  function handleCreate() {
    if (!form.jobName.trim()) { setError("Job name is required."); return; }
    if (!form.duration)       { setError("Project duration is required."); return; }
    onCreate({
      ...form,
      id: generateProjectId(),
      sqft: parseDollar(form.sqft),
      duration: parseDollar(form.duration),
      projectedFee: parseDollar(form.projectedFee),
      createdAt: new Date().toISOString().slice(0,10),
      status: "Bidding",
    });
  }

  return (
    <div style={s.overlay} onClick={e => e.target === e.currentTarget && onClose()}>
      <div style={{ backgroundColor: C.white, borderRadius: 4, width: 640, maxHeight: "90vh", boxShadow: "0 16px 48px rgba(0,0,0,0.25)", display: "flex", flexDirection: "column" }}>
        <div style={{ ...s.modalHeader, borderBottom: `1px solid ${C.border}` }}>
          <div>
            <div style={s.modalTitle}>New Project Estimate</div>
            <div style={s.modalSubtitle}>Create a project to begin estimating</div>
          </div>
          <button style={s.closeBtn} onClick={onClose}>×</button>
        </div>
        <div style={{ ...s.modalBody, overflowY: "auto" }}>
          {error && <div style={s.errorMsg}>{error}</div>}

          <div style={s.sectionDivider}>Project Info</div>
          <div style={s.grid2}>
            <div style={s.field}>
              <label style={s.label}>Job No.</label>
              <input style={s.input} value={form.jobNo} onChange={f("jobNo")} placeholder="e.g. MCG-2026-001" />
            </div>
            <div style={s.field}>
              <label style={s.label}>Project Type</label>
              <select style={s.fieldSelect} value={form.projectType} onChange={f("projectType")}>
                {["Commercial","Healthcare","Education","Industrial","Multifamily","Hospitality","Mixed-Use","Other"].map(t => <option key={t}>{t}</option>)}
              </select>
            </div>
          </div>
          <div style={s.field}>
            <label style={s.label}>Job Name *</label>
            <input style={s.input} value={form.jobName} onChange={f("jobName")} placeholder="e.g. ASU Gotaas Hall Renovation" autoFocus />
          </div>
          <div style={s.field}>
            <label style={s.label}>Location</label>
            <input style={s.input} value={form.location} onChange={f("location")} placeholder="City, ST" />
          </div>
          <div style={s.grid3}>
            <div style={s.field}>
              <label style={s.label}>Building Area (SF)</label>
              <input style={{ ...s.input, fontFamily: "monospace" }} type="number" min="0" value={form.sqft} onChange={f("sqft")} placeholder="0" />
            </div>
            <div style={s.field}>
              <label style={s.label}>Duration (months) *</label>
              <input style={{ ...s.input, fontFamily: "monospace" }} type="number" min="1" value={form.duration} onChange={f("duration")} placeholder="0" />
            </div>
            <div style={s.field}>
              <label style={s.label}>Estimator</label>
              <input style={s.input} value={form.estimator} onChange={f("estimator")} />
            </div>
          </div>

          <div style={s.sectionDivider}>Bid Info</div>
          <div style={s.grid3}>
            <div style={s.field}>
              <label style={s.label}>Industry</label>
              <select style={s.fieldSelect} value={form.industry} onChange={f("industry")}>
                {INDUSTRIES.map(i => <option key={i}>{i}</option>)}
              </select>
            </div>
            <div style={s.field}>
              <label style={s.label}>Bid Type</label>
              <select style={s.fieldSelect} value={form.bidType} onChange={f("bidType")}>
                {BID_TYPES.map(t => <option key={t}>{t}</option>)}
              </select>
            </div>
            <div style={s.field}>
              <label style={s.label}>Bid Due Date</label>
              <input style={s.input} type="date" value={form.dueDate} onChange={f("dueDate")} />
            </div>
          </div>
          <div style={s.grid2}>
            <div style={s.field}>
              <label style={s.label}>Client</label>
              <input style={s.input} value={form.client} onChange={f("client")} placeholder="Client company name" />
            </div>
            <div style={s.field}>
              <label style={s.label}>Projected Fee ($)</label>
              <input style={{ ...s.input, fontFamily: "monospace" }} type="number" min="0" value={form.projectedFee} onChange={f("projectedFee")} placeholder="0" />
            </div>
          </div>
          <div style={s.grid2}>
            <div style={s.field}>
              <label style={s.label}>Contact Name</label>
              <input style={s.input} value={form.contactName} onChange={f("contactName")} placeholder="Primary contact" />
            </div>
            <div style={s.field}>
              <label style={s.label}>Contact Email</label>
              <input style={s.input} type="email" value={form.contactEmail} onChange={f("contactEmail")} placeholder="email@company.com" />
            </div>
          </div>
        </div>
        <div style={{ ...s.modalFooter, borderTop: `1px solid ${C.border}`, justifyContent: "flex-end", gap: 8 }}>
          <button style={s.btn()} onClick={onClose}>Cancel</button>
          <button style={s.btn("primary")} onClick={handleCreate}>Create Estimate →</button>
        </div>
      </div>
    </div>
  );
}

// ── Edit Project Modal ───────────────────────────────────────────────────────
function EditProjectModal({ project, onClose, onSave }) {
  const [form, setForm] = useState({
    jobNo:        project.jobNo        || "",
    jobName:      project.jobName      || "",
    location:     project.location     || "",
    sqft:         project.sqft         || "",
    estimator:    project.estimator    || "",
    duration:     project.duration     || "",
    projectType:  project.projectType  || "Commercial",
    client:       project.client       || "",
    industry:     project.industry     || "Retail",
    bidType:      project.bidType      || "Hard Bid",
    projectedFee: project.projectedFee || "",
    dueDate:      project.dueDate      || "",
    contactName:  project.contactName  || "",
    contactEmail: project.contactEmail || "",
    createdAt:    project.createdAt    || "",
  });
  const [error, setError] = useState("");
  const f = k => e => setForm(p => ({ ...p, [k]: e.target.value }));

  function handleSave() {
    if (!form.jobName.trim()) { setError("Job name is required."); return; }
    if (!form.duration)       { setError("Duration is required."); return; }
    onSave({
      ...project,
      ...form,
      sqft:         parseDollar(form.sqft),
      duration:     parseDollar(form.duration),
      projectedFee: parseDollar(form.projectedFee),
    });
    onClose();
  }

  return (
    <div style={s.overlay} onClick={e => e.target === e.currentTarget && onClose()}>
      <div style={{ backgroundColor: C.white, borderRadius: 4, width: 640, maxHeight: "90vh", boxShadow: "0 16px 48px rgba(0,0,0,0.25)", display: "flex", flexDirection: "column" }}>
        <div style={{ ...s.modalHeader, borderBottom: `1px solid ${C.border}` }}>
          <div>
            <div style={s.modalTitle}>Edit Project</div>
            <div style={s.modalSubtitle}>{project.jobNo} — {project.jobName}</div>
          </div>
          <button style={s.closeBtn} onClick={onClose}>×</button>
        </div>
        <div style={{ ...s.modalBody, overflowY: "auto" }}>
          {error && <div style={s.errorMsg}>{error}</div>}

          <div style={s.sectionDivider}>Project Info</div>
          <div style={s.grid2}>
            <div style={s.field}>
              <label style={s.label}>Job No.</label>
              <input style={s.input} value={form.jobNo} onChange={f("jobNo")} />
            </div>
            <div style={s.field}>
              <label style={s.label}>Project Type</label>
              <select style={s.fieldSelect} value={form.projectType} onChange={f("projectType")}>
                {["Commercial","Healthcare","Education","Industrial","Multifamily","Hospitality","Mixed-Use","Other"].map(t => <option key={t}>{t}</option>)}
              </select>
            </div>
          </div>
          <div style={s.field}>
            <label style={s.label}>Job Name *</label>
            <input style={s.input} value={form.jobName} onChange={f("jobName")} autoFocus />
          </div>
          <div style={s.field}>
            <label style={s.label}>Location</label>
            <input style={s.input} value={form.location} onChange={f("location")} placeholder="City, ST" />
          </div>
          <div style={s.grid3}>
            <div style={s.field}>
              <label style={s.label}>Building Area (SF)</label>
              <input style={{ ...s.input, fontFamily: "monospace" }} type="number" min="0" value={form.sqft} onChange={f("sqft")} />
            </div>
            <div style={s.field}>
              <label style={s.label}>Duration (months) *</label>
              <input style={{ ...s.input, fontFamily: "monospace" }} type="number" min="1" value={form.duration} onChange={f("duration")} />
            </div>
            <div style={s.field}>
              <label style={s.label}>Estimator</label>
              <input style={s.input} value={form.estimator} onChange={f("estimator")} />
            </div>
          </div>
          <div style={s.grid2}>
            <div style={s.field}>
              <label style={s.label}>Created Date</label>
              <input style={s.input} type="date" value={form.createdAt} onChange={f("createdAt")} />
            </div>
          </div>

          <div style={s.sectionDivider}>Bid Info</div>
          <div style={s.grid3}>
            <div style={s.field}>
              <label style={s.label}>Industry</label>
              <select style={s.fieldSelect} value={form.industry} onChange={f("industry")}>
                {INDUSTRIES.map(i => <option key={i}>{i}</option>)}
              </select>
            </div>
            <div style={s.field}>
              <label style={s.label}>Bid Type</label>
              <select style={s.fieldSelect} value={form.bidType} onChange={f("bidType")}>
                {BID_TYPES.map(t => <option key={t}>{t}</option>)}
              </select>
            </div>
            <div style={s.field}>
              <label style={s.label}>Bid Due Date</label>
              <input style={s.input} type="date" value={form.dueDate} onChange={f("dueDate")} />
            </div>
          </div>
          <div style={s.grid2}>
            <div style={s.field}>
              <label style={s.label}>Client</label>
              <input style={s.input} value={form.client} onChange={f("client")} placeholder="Client company name" />
            </div>
            <div style={s.field}>
              <label style={s.label}>Projected Fee ($)</label>
              <input style={{ ...s.input, fontFamily: "monospace" }} type="number" min="0" value={form.projectedFee} onChange={f("projectedFee")} />
            </div>
          </div>
          <div style={s.grid2}>
            <div style={s.field}>
              <label style={s.label}>Contact Name</label>
              <input style={s.input} value={form.contactName} onChange={f("contactName")} />
            </div>
            <div style={s.field}>
              <label style={s.label}>Contact Email</label>
              <input style={s.input} type="email" value={form.contactEmail} onChange={f("contactEmail")} />
            </div>
          </div>
        </div>
        <div style={{ ...s.modalFooter, borderTop: `1px solid ${C.border}`, justifyContent: "flex-end", gap: 8 }}>
          <button style={s.btn()} onClick={onClose}>Cancel</button>
          <button style={s.btn("primary")} onClick={handleSave}>Save Changes</button>
        </div>
      </div>
    </div>
  );
}

// ── Project List ──────────────────────────────────────────────────────────────
const BID_STATUSES = ["Bidding", "Pending", "Won", "Lost", "No Bid"];

const STATUS_STYLES = {
  "Bidding": { backgroundColor: "#DBEAFE", color: "#1D4ED8" },
  "Pending": { backgroundColor: "#FEF3C7", color: "#92400E" },
  "Won":     { backgroundColor: "#D1FAE5", color: "#065F46" },
  "Lost":    { backgroundColor: "#FEE2E2", color: "#991B1B" },
  "No Bid":  { backgroundColor: "#F1F5F9", color: "#475569" },
  "Active":  { backgroundColor: "#D1FAE5", color: "#065F46" },
};

function ProjectList({ projects, onSelect, onNew, onStatusChange, onEdit }) {
  return (
    <div style={{ flex: 1, display: "flex", flexDirection: "column", overflow: "hidden" }}>
      <div style={{ padding: "16px 24px 12px", borderBottom: `1px solid ${C.border}`, display: "flex", alignItems: "center", justifyContent: "space-between", flexShrink: 0 }}>
        <div>
          <div style={{ fontSize: 15, fontWeight: 700, color: C.textPrimary }}>Project Estimates</div>
          <div style={{ fontSize: 12, color: C.textMuted, marginTop: 2 }}>{projects.length} project{projects.length !== 1 ? "s" : ""}</div>
        </div>
        <button style={s.btn("primary")} onClick={onNew}>+ New Estimate</button>
      </div>

      {projects.length === 0 ? (
        <div style={{ flex: 1, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", color: C.textMuted, gap: 12 }}>
          <div style={{ fontSize: 36 }}>📋</div>
          <div style={{ fontWeight: 600, fontSize: 14, color: C.textPrimary }}>No estimates yet</div>
          <div style={{ fontSize: 13 }}>Create your first project estimate to get started.</div>
          <button style={{ ...s.btn("primary"), marginTop: 8 }} onClick={onNew}>+ New Estimate</button>
        </div>
      ) : (
        <div style={{ overflowY: "auto", flex: 1 }}>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr>
                {["Job No.", "Job Name", "Type", "Location", "Area (SF)", "Duration", "Estimator", "Created", "Value", "Status", "", ""].map(h => (
                  <th key={h} style={{ padding: "7px 14px", fontSize: 11, fontWeight: 700, color: C.textMuted, textTransform: "uppercase", letterSpacing: "0.06em", borderBottom: `2px solid ${C.border}`, textAlign: "left", backgroundColor: C.contentBg, whiteSpace: "nowrap" }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {projects.map((p, i) => (
                <tr key={p.id} style={{ backgroundColor: i % 2 === 0 ? C.white : C.tableStripe, cursor: "pointer" }} onClick={() => onSelect(p)}>
                  <td style={{ padding: "8px 14px", fontSize: 12, fontFamily: "monospace", color: C.accentLight, borderBottom: `1px solid ${C.border}` }}>{p.jobNo}</td>
                  <td style={{ padding: "8px 14px", fontSize: 12, fontWeight: 600, color: C.textPrimary, borderBottom: `1px solid ${C.border}` }}>{p.jobName}</td>
                  <td style={{ padding: "8px 14px", fontSize: 12, color: C.textMuted, borderBottom: `1px solid ${C.border}` }}>{p.projectType}</td>
                  <td style={{ padding: "8px 14px", fontSize: 12, color: C.textMuted, borderBottom: `1px solid ${C.border}` }}>{p.location || "—"}</td>
                  <td style={{ padding: "8px 14px", fontSize: 12, fontFamily: "monospace", color: C.textMuted, borderBottom: `1px solid ${C.border}` }}>{p.sqft ? Number(p.sqft).toLocaleString() : "—"}</td>
                  <td style={{ padding: "8px 14px", fontSize: 12, color: C.textMuted, borderBottom: `1px solid ${C.border}` }}>{p.duration} mo.</td>
                  <td style={{ padding: "8px 14px", fontSize: 12, color: C.textMuted, borderBottom: `1px solid ${C.border}` }}>{p.estimator}</td>
                  <td style={{ padding: "8px 14px", fontSize: 12, color: C.textMuted, borderBottom: `1px solid ${C.border}` }}>{p.createdAt}</td>
                  <td style={{ padding: "8px 14px", fontSize: 12, fontFamily: "monospace", fontWeight: 600, color: C.textPrimary, borderBottom: `1px solid ${C.border}`, whiteSpace: "nowrap" }}>
                    {p.projectTotal > 0 ? ("$" + Number(p.projectTotal).toLocaleString("en-US", { minimumFractionDigits: 0, maximumFractionDigits: 0 })) : <span style={{ color: C.textLight }}>—</span>}
                  </td>
                  <td style={{ padding: "6px 14px", borderBottom: `1px solid ${C.border}` }} onClick={e => e.stopPropagation()}>
                    <select
                      value={p.status}
                      onChange={e => onStatusChange(p.id, e.target.value)}
                      style={{
                        padding: "2px 6px", borderRadius: 3, fontSize: 11, fontWeight: 600, cursor: "pointer",
                        border: "1px solid transparent", outline: "none",
                        ...(STATUS_STYLES[p.status] || STATUS_STYLES["Bidding"]),
                      }}
                    >
                      {BID_STATUSES.map(st => <option key={st} value={st}>{st}</option>)}
                    </select>
                  </td>
                  <td style={{ padding: "6px 10px", fontSize: 12, borderBottom: `1px solid ${C.border}`, whiteSpace: "nowrap" }} onClick={e => e.stopPropagation()}>
                    <button onClick={() => onEdit(p)} style={{ ...s.btn(), fontSize: 11, padding: "2px 8px" }}>Edit</button>
                  </td>
                  <td style={{ padding: "8px 14px", fontSize: 12, color: C.accentLight, borderBottom: `1px solid ${C.border}`, whiteSpace: "nowrap" }}>Open →</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

// ── Estimate Workspace ────────────────────────────────────────────────────────
function EstimateWorkspace({ project, onBack, onDelete, onProjectUpdate, firmRates, firmDefaults, onFirmRatesChange, onFirmDefaultsChange }) {
  const TABS = [
    { key: "recap",       label: "Recap"             },
    { key: "bidsheet",    label: "Bid Sheet"          },
    { key: "supervision", label: "Supervision"        },
    { key: "gc",          label: "General Conditions" },
    { key: "equipment",   label: "Equipment"          },
    { key: "insurance",   label: "Insurance & Permits"},
    { key: "allowances",  label: "Allowances"         },
    { key: "sov",         label: "SOV"                },
  ];

  const [activeTab,       setActiveTab]       = useState("recap");
  const [saveStatus,      setSaveStatus]      = useState("saved"); // "saved" | "saving"

  // Load persisted estimate data for this project
  const _saved = useMemo(() => lsGet(lsEstKey(project.id), {}), [project.id]);

  const [supervisionData, setSupervisionData] = useState(() => _saved.supervisionData || {});
  const [sheet2500,       setSheet2500]       = useState(() => _saved.sheet2500 || {});
  const [sheet2600,       setSheet2600]       = useState(() => _saved.sheet2600 || {});
  const [sheet2700,       setSheet2700]       = useState(() => _saved.sheet2700 || {});
  const [sheet2900,       setSheet2900]       = useState(() => _saved.sheet2900 || {});
  const [bidData,         setBidData]         = useState(() => _saved.bidData || {});
  const [recapData,       setRecapData]       = useState(() => _saved.recapData || {});
  const [insSsRate,       setInsSsRate]       = useState(() => _saved.insSsRate || String((FIRM_DEFAULTS.insSsRate * 100).toFixed(2)));
  const [assumptions,     setAssumptions]     = useState(() => _saved.assumptions || SUPERVISION_ASSUMPTION_DEFAULTS);
  const [showSettings,    setShowSettings]    = useState(false);
  const [showDeleteConfirm, setShowDeleteConfirm] = useState(false);

  // Auto-save whenever any estimate data changes
  useEffect(() => {
    setSaveStatus("saving");
    const timer = setTimeout(() => {
      const _total = (() => {
        try {
          const fR = parseDollar(recapData.__feeRate  ?? (FIRM_DEFAULTS.feeRate*100));
          const bR = parseDollar(recapData.__bondRate ?? (FIRM_DEFAULTS.bondRate*100));
          const gR = parseDollar(recapData.__glRate   ?? (FIRM_DEFAULTS.glRate*100));
          const tR = parseDollar(recapData.__grtRate  ?? (FIRM_DEFAULTS.grtRate*100));
          const cA = parseDollar(recapData.__contingencyAmt ?? 0);
          const iR = parseDollar(insSsRate ?? (FIRM_DEFAULTS.insSsRate*100));
          const bidByCode = {};
          BIDSHEET_TRADES.forEach((trade, idx) => {
            const d = (bidData||{})[idx] || {};
            const amt = parseDollar(d.amount ?? 0); if (!amt) return;
            const type = d.type ?? trade.type;
            const k5 = trade.code.slice(0,5), k2 = trade.code.slice(0,2);
            const rk = RECAP_TRADE_ROWS.find(r=>r.code===k5)?.code || RECAP_TRADE_ROWS.find(r=>r.code.slice(0,2)===k2)?.code || null;
            if (!rk) return;
            if (!bidByCode[rk]) bidByCode[rk] = { material:0, sub:0 };
            if (type==="M") bidByCode[rk].material += amt;
            else if (type==="S") bidByCode[rk].sub += amt;
          });
          let tradeSum=0, laborSum=0;
          RECAP_TRADE_ROWS.forEach(row => {
            const d = recapData[row.code]||{}, bid = bidByCode[row.code]||{};
            const l = parseDollar(d.labor!==undefined&&d.labor!==""?d.labor:0);
            const m = parseDollar(d.material!==undefined&&d.material!==""?d.material:(bid.material??0));
            const sb= parseDollar(d.sub!==undefined&&d.sub!==""?d.sub:(bid.sub??0));
            tradeSum += l+m+sb; laborSum += l;
          });
          const gcT = (sheetTotals["2500"]||0)+(sheetTotals["2600"]||0)+(sheetTotals["2700"]||0)+(sheetTotals["2900"]||0);
          const sub = tradeSum+gcT;
          const ins = laborSum*(iR/100);
          const s2  = sub+ins;
          const fee = s2*(fR/100), bond=(s2+fee)*(bR/100), gl=(s2+fee)*(gR/100), grt=(s2+fee)*(tR/100);
          return s2+fee+bond+gl+grt+cA;
        } catch { return 0; }
      })();
      lsSet(lsEstKey(project.id), {
        supervisionData, sheet2500, sheet2600, sheet2700, sheet2900,
        bidData, recapData, insSsRate, assumptions, projectTotal: _total,
      });
      if (onProjectUpdate) onProjectUpdate({ projectTotal: _total });
      setSaveStatus("saved");
    }, 600);
    return () => clearTimeout(timer);
  }, [supervisionData, sheet2500, sheet2600, sheet2700, sheet2900, bidData, recapData, insSsRate, assumptions]);

  // Calc supervision totals to pass down into 2500 auto-calc rows
  const supervisionTotals = useMemo(() => {
    const assumptions = { hoursPerWeek: 40, weeksPerMonth: 4.33, mileageRate: 0.725 };
    let totalSalary = 0, perDiem = 0, travel = 0, fuelTotal = 0, safetyDir = 0;
    SUPERVISION_ROLES.forEach(role => {
      const d = supervisionData[role.key] || {};
      const hourlyRate = firmRates[role.key] ?? role.defaultRate;
      const weeks  = parseDollar(d.weeks ?? 0);
      const salary = weeks * hourlyRate * assumptions.hoursPerWeek;
      const trips  = parseDollar(d.trips ?? 0);
      const milesPerTrip = parseDollar(d.milesPerTrip ?? 155);
      const mileageCost  = trips * milesPerTrip * assumptions.mileageRate;
      const perDiemMo    = parseDollar(d.perDiem ?? 0);
      const perDiemAmt   = perDiemMo * (weeks / assumptions.weeksPerMonth);
      if (role.key === "safety_director") safetyDir = salary + mileageCost + perDiemAmt;
      totalSalary += salary;
      perDiem     += perDiemAmt;
      travel      += mileageCost;
    });
    return { supervision: totalSalary, perDiem, travel, fuel: fuelTotal, safetyDir };
  }, [supervisionData, firmRates]);

  // Live sheet totals for RECAP
  const sheetTotals = useMemo(() => {
    const calcSheet = (templateRows, sheetData) => templateRows.reduce((acc, row, idx) => {
      if (row.isSection) return acc;
      if (row.readOnly && row.autoCalc) return acc + (supervisionTotals[row.autoCalc] || 0);
      const d = sheetData[idx] || {};
      return acc + parseDollar(d.qty ?? 0) * parseDollar(d.sub_up !== undefined ? d.sub_up : (row.sub_up ?? 0));
    }, 0);
    return {
      "2500": calcSheet(SHEET_2500_ROWS, sheet2500) + supervisionTotals.supervision,
      "2600": calcSheet(SHEET_2600_ROWS, sheet2600),
      "2700": calcSheet(SHEET_2700_ROWS, sheet2700),
      "2900": calcSheet(SHEET_2900_ROWS, sheet2900),
    };
  }, [sheet2500, sheet2600, sheet2700, sheet2900, supervisionTotals]);

  const durationWeeks = Math.round((parseDollar(project.duration) || 0) * 4.33);

  // Compute live project total and push to parent for display in project list
  const projectTotal = useMemo(() => {
    const feeRate        = parseDollar(recapData.__feeRate        ?? (firmDefaults.feeRate * 100));
    const bondRate       = parseDollar(recapData.__bondRate       ?? (firmDefaults.bondRate * 100));
    const glRate         = parseDollar(recapData.__glRate         ?? (firmDefaults.glRate * 100));
    const grtRate        = parseDollar(recapData.__grtRate        ?? (firmDefaults.grtRate * 100));
    const contingencyAmt = parseDollar(recapData.__contingencyAmt ?? 0);
    const insSsRateVal   = parseDollar(insSsRate ?? (firmDefaults.insSsRate * 100));
    const bidByCode = {};
    BIDSHEET_TRADES.forEach((trade, idx) => {
      const d = (bidData||{})[idx] || {};
      const amount = parseDollar(d.amount ?? 0);
      if (!amount) return;
      const type = d.type ?? trade.type;
      const key5 = trade.code.slice(0,5);
      const key2 = trade.code.slice(0,2);
      const rk = RECAP_TRADE_ROWS.find(r => r.code === key5)?.code
              || RECAP_TRADE_ROWS.find(r => r.code.slice(0,2) === key2)?.code || null;
      if (!rk) return;
      if (!bidByCode[rk]) bidByCode[rk] = { material: 0, sub: 0 };
      if (type === "M") bidByCode[rk].material += amount;
      else if (type === "S") bidByCode[rk].sub += amount;
    });
    const tradeLines = RECAP_TRADE_ROWS.map(row => {
      const d = recapData[row.code] || {};
      const bid = bidByCode[row.code] || {};
      const labor    = parseDollar(d.labor    !== undefined && d.labor    !== "" ? d.labor    : 0);
      const material = parseDollar(d.material !== undefined && d.material !== "" ? d.material : (bid.material ?? 0));
      const sub      = parseDollar(d.sub      !== undefined && d.sub      !== "" ? d.sub      : (bid.sub ?? 0));
      return labor + material + sub;
    });
    const gcTotal   = (sheetTotals["2500"]||0)+(sheetTotals["2600"]||0)+(sheetTotals["2700"]||0)+(sheetTotals["2900"]||0);
    const subtotal  = tradeLines.reduce((a,v)=>a+v,0) + gcTotal;
    const laborSum  = RECAP_TRADE_ROWS.reduce((a, row) => {
      const d = recapData[row.code] || {};
      return a + parseDollar(d.labor !== undefined && d.labor !== "" ? d.labor : 0);
    }, 0);
    const insSsAmt  = laborSum * (insSsRateVal / 100);
    const sub2      = subtotal + insSsAmt;
    const feeAmt    = sub2 * (feeRate/100);
    const bondAmt   = (sub2+feeAmt) * (bondRate/100);
    const glAmt     = (sub2+feeAmt) * (glRate/100);
    const grtAmt    = (sub2+feeAmt) * (grtRate/100);
    return sub2 + feeAmt + bondAmt + glAmt + grtAmt + contingencyAmt;
  }, [recapData, bidData, sheetTotals, insSsRate, firmDefaults]);



  // ── Excel Export ────────────────────────────────────────────────────────────
  function handleExportExcel() {
    const wb = XLSX.utils.book_new();
    const sqft = parseDollar(project.sqft || 0);

    // ── helpers ──────────────────────────────────────────────────────────────
    const dollarFmt = '#,##0.00';
    const pctFmt    = '0.00"%"';
    const hdrStyle  = { font: { bold: true, color: { rgb: "FFFFFF" } }, fill: { fgColor: { rgb: "0F1923" } }, alignment: { horizontal: "center" } };
    const subHdr    = { font: { bold: true }, fill: { fgColor: { rgb: "E2E8F0" } } };
    const totalRow  = { font: { bold: true, sz: 12 }, fill: { fgColor: { rgb: "0F1923" } }, font2: { color: { rgb: "FFFFFF" } } };

    function addSheet(name, rows) {
      const ws = XLSX.utils.aoa_to_sheet(rows);
      XLSX.utils.book_append_sheet(wb, ws, name);
      return ws;
    }

    function setColWidths(ws, widths) {
      ws['!cols'] = widths.map(w => ({ wch: w }));
    }

    // ── Shared markup calcs (mirrors RecapSheet / SOVSheet) ──────────────────
    const feeRate        = parseDollar(recapData.__feeRate        ?? (firmDefaults.feeRate * 100));
    const bondRate       = parseDollar(recapData.__bondRate       ?? (firmDefaults.bondRate * 100));
    const glRate         = parseDollar(recapData.__glRate         ?? (firmDefaults.glRate * 100));
    const grtRate        = parseDollar(recapData.__grtRate        ?? (firmDefaults.grtRate * 100));
    const contingencyAmt = parseDollar(recapData.__contingencyAmt ?? 0);
    const insSsRateVal   = parseDollar(insSsRate ?? (firmDefaults.insSsRate * 100));

    const bidByCode = {};
    BIDSHEET_TRADES.forEach((trade, idx) => {
      const d = (bidData || {})[idx] || {};
      const amount = parseDollar(d.amount ?? 0);
      if (!amount) return;
      const type = d.type ?? trade.type;
      const key5 = trade.code.slice(0, 5);
      const key2 = trade.code.slice(0, 2);
      const rk   = RECAP_TRADE_ROWS.find(r => r.code === key5)?.code
                || RECAP_TRADE_ROWS.find(r => r.code.slice(0,2) === key2)?.code || null;
      if (!rk) return;
      if (!bidByCode[rk]) bidByCode[rk] = { material: 0, sub: 0 };
      if (type === "M") bidByCode[rk].material += amount;
      else if (type === "S") bidByCode[rk].sub += amount;
    });

    const tradeLines = RECAP_TRADE_ROWS.map(row => {
      const d = recapData[row.code] || {};
      const bid = bidByCode[row.code] || {};
      const labor    = parseDollar(d.labor    !== undefined && d.labor    !== "" ? d.labor    : 0);
      const material = parseDollar(d.material !== undefined && d.material !== "" ? d.material : (bid.material ?? 0));
      const sub      = parseDollar(d.sub      !== undefined && d.sub      !== "" ? d.sub      : (bid.sub      ?? 0));
      return { ...row, labor, material, sub, total: labor + material + sub };
    });

    const gcTotal    = (sheetTotals["2500"]||0)+(sheetTotals["2600"]||0)+(sheetTotals["2700"]||0)+(sheetTotals["2900"]||0);
    const tradeTotal = tradeLines.reduce((a,r) => a+r.total, 0);
    const subtotal   = tradeTotal + gcTotal;
    const insSsAmt   = tradeLines.reduce((a,r) => a+r.labor, 0) * (insSsRateVal/100);
    const subtotal2  = subtotal + insSsAmt;
    const feeAmt     = subtotal2 * (feeRate/100);
    const bondAmt    = (subtotal2+feeAmt) * (bondRate/100);
    const glAmt      = (subtotal2+feeAmt) * (glRate/100);
    const grtAmt     = (subtotal2+feeAmt) * (grtRate/100);
    const projectTotal = subtotal2 + feeAmt + bondAmt + glAmt + grtAmt + contingencyAmt;

    // ── 1. SOV sheet ─────────────────────────────────────────────────────────
    const today = new Date().toLocaleDateString("en-US");
    const sovRows = [
      ["MERIDIAN CONSTRUCTION GROUP"],
      ["SCHEDULE OF VALUES — PROPOSAL"],
      [],
      ["Project:", project.jobName,       "",  "Job No.:",    project.jobNo],
      ["Location:", project.location||"—","",  "Type:",       project.projectType],
      ["Estimator:", project.estimator,   "",  "Date:",       today],
      ["Building Area:", sqft > 0 ? sqft : "—", "SF", "Duration:", project.duration + " months"],
      [],
      ["CSI CODE", "DESCRIPTION", "TOTAL", "$/SF"],
      ...tradeLines.filter(r => r.total > 0).map(r => [
        r.code, r.label, r.total, sqft > 0 ? r.total/sqft : ""
      ]),
      gcTotal > 0 ? ["01 00 00", "GENERAL CONDITIONS", gcTotal, sqft > 0 ? gcTotal/sqft : ""] : null,
      [],
      ["", "SUBTOTAL", subtotal2, sqft > 0 ? subtotal2/sqft : ""],
      feeAmt  > 0 ? ["", `GC Fee (${feeRate.toFixed(2)}%)`,                    feeAmt,  sqft > 0 ? feeAmt/sqft  : ""] : null,
      bondAmt > 0 ? ["", `Performance & Payment Bond (${bondRate.toFixed(2)}%)`, bondAmt, sqft > 0 ? bondAmt/sqft : ""] : null,
      glAmt   > 0 ? ["", `General Liability Insurance (${glRate.toFixed(2)}%)`,  glAmt,   sqft > 0 ? glAmt/sqft   : ""] : null,
      grtAmt  > 0 ? ["", `Gross Receipts Tax (${grtRate.toFixed(2)}%)`,          grtAmt,  sqft > 0 ? grtAmt/sqft  : ""] : null,
      contingencyAmt > 0 ? ["", "Contingency", contingencyAmt, sqft > 0 ? contingencyAmt/sqft : ""] : null,
      [],
      ["", "TOTAL PROPOSED CONTRACT SUM", projectTotal, sqft > 0 ? projectTotal/sqft : ""],
    ].filter(Boolean);

    const sovWs = addSheet("SOV", sovRows);
    setColWidths(sovWs, [12, 44, 16, 12]);

    // ── 2. RECAP sheet ────────────────────────────────────────────────────────
    const recapRows = [
      ["PROJECT RECAP — " + project.jobName],
      [],
      ["CODE", "DESCRIPTION", "LABOR", "MATERIAL", "SUB / EQUIP", "TOTAL", "$/SF"],
      ...tradeLines.map(r => [r.code, r.label, r.labor||"", r.material||"", r.sub||"", r.total||"", sqft > 0 && r.total ? r.total/sqft : ""]),
      [],
      ["2500", "GENERAL CONDITIONS",           "", "", sheetTotals["2500"]||"", sheetTotals["2500"]||"", sqft > 0 && sheetTotals["2500"] ? sheetTotals["2500"]/sqft : ""],
      ["2600", "EQUIPMENT",                    "", "", sheetTotals["2600"]||"", sheetTotals["2600"]||"", sqft > 0 && sheetTotals["2600"] ? sheetTotals["2600"]/sqft : ""],
      ["2700", "INSURANCE, PERMITS & BONDING", "", "", sheetTotals["2700"]||"", sheetTotals["2700"]||"", sqft > 0 && sheetTotals["2700"] ? sheetTotals["2700"]/sqft : ""],
      ["2900", "ALLOWANCES / TESTING",         "", "", sheetTotals["2900"]||"", sheetTotals["2900"]||"", sqft > 0 && sheetTotals["2900"] ? sheetTotals["2900"]/sqft : ""],
      [],
      ["", "SUBTOTAL", tradeLines.reduce((a,r)=>a+r.labor,0), tradeLines.reduce((a,r)=>a+r.material,0), tradeLines.reduce((a,r)=>a+r.sub,0)+gcTotal, subtotal, sqft > 0 ? subtotal/sqft : ""],
      ["", `GC Fee (${feeRate.toFixed(2)}%)`,                    "", "", "", feeAmt,  sqft > 0 ? feeAmt/sqft  : ""],
      ["", `Bond (${bondRate.toFixed(2)}%)`,                     "", "", "", bondAmt, sqft > 0 ? bondAmt/sqft : ""],
      ["", `GL Insurance (${glRate.toFixed(2)}%)`,               "", "", "", glAmt,   sqft > 0 ? glAmt/sqft   : ""],
      ["", `Gross Receipts Tax (${grtRate.toFixed(2)}%)`,        "", "", "", grtAmt,  sqft > 0 ? grtAmt/sqft  : ""],
      contingencyAmt > 0 ? ["", "Contingency",                  "", "", "", contingencyAmt, sqft > 0 ? contingencyAmt/sqft : ""] : null,
      [],
      ["", "PROJECT TOTAL", "", "", "", projectTotal, sqft > 0 ? projectTotal/sqft : ""],
    ].filter(Boolean);

    const recapWs = addSheet("RECAP", recapRows);
    setColWidths(recapWs, [10, 40, 16, 16, 16, 16, 10]);

    // ── 3. Bid Sheet ─────────────────────────────────────────────────────────
    const bidRows = [
      ["BID SHEET — " + project.jobName],
      [],
      ["CSI CODE", "ITEM", "AMOUNT", "SUB / VENDOR", "TYPE", "SUBGUARD %"],
      ...BIDSHEET_TRADES.map((trade, idx) => {
        const d = (bidData||{})[idx] || {};
        return [trade.code, trade.item, parseDollar(d.amount??0)||"", d.vendor||"", d.type??trade.type, d.subguardPct||""];
      }),
      [],
      ["", "TOTAL BID", BIDSHEET_TRADES.reduce((a,_,idx) => a + parseDollar(((bidData||{})[idx]||{}).amount??0), 0), "", "", ""],
    ];
    const bidWs = addSheet("Bid Sheet", bidRows);
    setColWidths(bidWs, [10, 44, 16, 28, 8, 10]);

    // ── 4. Supervision sheet ─────────────────────────────────────────────────
    const A = assumptions;
    const supRows = [
      ["SUPERVISION & ADMINISTRATION — " + project.jobName],
      [`Assumptions: ${A.hoursPerWeek} hrs/wk | ${A.weeksPerMonth} wks/mo | $${A.mileageRate}/mi | Hotel $${A.hotelRate}/night | Plane $${A.planeRate} | Car Rental $${A.carRentalRate}/day`],
      [],
      ["ROLE", "HRLY RATE", "WEEKS", "SALARY $", "FLEET VEH.", "DAILY MI.", "FUEL $", "COMPUTER/MO", "CELL/MO", "OTHER/MO", "PER DIEM", "PER DIEM $", "TRIPS (MILEAGE)", "MI/TRIP", "MILEAGE $", "TRIPS (TRAVEL)", "DAYS/TRIP", "HOTEL", "PLANE", "CAR RENTAL", "TRAVEL $", "TOTAL COST", "LOADED $/HR"],
      ...SUPERVISION_ROLES.map(role => {
        const d = supervisionData[role.key] || {};
        const hourlyRate = firmRates[role.key] ?? role.defaultRate;
        const weeks  = parseDollar(d.weeks??0);
        const salary = weeks * hourlyRate * A.hoursPerWeek;
        const months = weeks / A.weeksPerMonth;
        const hasFleet    = d.fleetVehicle === "Yes";
        const dailyMiles  = parseDollar(d.dailyMiles ?? A.defaultDailyMiles);
        const fleetFuel   = hasFleet ? weeks*5*dailyMiles/A.mpg*A.fuelPerGal : 0;
        const compAmt     = parseDollar(d.computerMo??0)*months;
        const cellAmt     = parseDollar(d.cellMo??0)*months;
        const otherAmt    = parseDollar(d.otherEquipMo??0)*months;
        const hasPerDiem  = d.perDiemYN === "Yes";
        const perDiemAmt  = hasPerDiem ? parseDollar(d.perDiemMo??0)*months : 0;
        const mTrips      = parseDollar(d.mileageTrips??0);
        const mMiles      = parseDollar(d.milesPerTrip??155);
        const mileageCost = mTrips * mMiles * A.mileageRate;
        const tTrips      = parseDollar(d.travelTrips??0);
        const tDays       = parseDollar(d.daysPerTrip??0);
        const hotelCost   = d.hotelRequired==="Yes" ? tTrips*tDays*A.hotelRate : 0;
        const planeCost   = d.planeTicket==="Yes"   ? tTrips*A.planeRate : 0;
        const carCost     = d.carRental==="Yes"     ? tTrips*tDays*A.carRentalRate : 0;
        const mealsCost   = tTrips*tDays*A.foodAllowance;
        const travelTotal = hotelCost+planeCost+carCost+mealsCost;
        const total       = salary+fleetFuel+compAmt+cellAmt+otherAmt+perDiemAmt+mileageCost+travelTotal;
        const loadedRate  = weeks > 0 ? total/(weeks*A.hoursPerWeek) : 0;
        return [
          role.label, hourlyRate, weeks||"", salary||"",
          d.fleetVehicle||"No", dailyMiles||"", fleetFuel||"",
          parseDollar(d.computerMo??0)||"", parseDollar(d.cellMo??0)||"", parseDollar(d.otherEquipMo??0)||"",
          d.perDiemYN||"No", perDiemAmt||"",
          mTrips||"", mMiles||"", mileageCost||"",
          tTrips||"", tDays||"", d.hotelRequired||"No", d.planeTicket||"No", d.carRental||"No",
          travelTotal||"", total||"", loadedRate||"",
        ];
      }),
      [],
      ["TOTAL", "", "", supervisionTotals.supervision||"", "", "", "", "", "", "", "", supervisionTotals.perDiem||"", "", "", supervisionTotals.mileage||"", "", "", "", "", "", supervisionTotals.travel||"", supervisionTotals.total||"", ""],
    ];
    const supWs = addSheet("Supervision", supRows);
    setColWidths(supWs, [28,10,8,14,10,10,10,12,10,10,10,12,14,8,12,14,10,8,8,10,12,14,12]);

    // ── 5. Detail sheets (GC/Equipment/Insurance/Allowances) ─────────────────
    const detailSheetDefs = [
      { name: "General Conditions", template: SHEET_2500_ROWS, data: sheet2500 },
      { name: "Equipment",          template: SHEET_2600_ROWS, data: sheet2600 },
      { name: "Insurance & Permits",template: SHEET_2700_ROWS, data: sheet2700 },
      { name: "Allowances",         template: SHEET_2900_ROWS, data: sheet2900 },
    ];

    detailSheetDefs.forEach(({ name, template, data }) => {
      const rows = [
        [name.toUpperCase() + " — " + project.jobName],
        [],
        ["CODE", "DESCRIPTION", "UNIT", "QTY", "UNIT PRICE", "TOTAL", "COMMENTS"],
        ...template.map((row, idx) => {
          if (row.isSection) return [row.section, "", "", "", "", "", ""];
          const d = data[idx] || {};
          if (row.readOnly && row.autoCalc) {
            const val = supervisionTotals[row.autoCalc] || 0;
            return [row.code||"", row.desc, row.unit||"", "", "", val||"", row.comments||""];
          }
          const qty    = parseDollar(d.qty    ?? 0);
          const sub_up = parseDollar(d.sub_up !== undefined ? d.sub_up : (row.sub_up ?? 0));
          const total  = qty * sub_up;
          return [row.code||"", d.desc !== undefined ? d.desc : (row.desc||""), row.unit||"", qty||"", sub_up||"", total||"", d.comments ?? row.comments ?? ""];
        }),
        [],
        ["", "TOTAL", "", "", "", template.reduce((acc, row, idx) => {
          if (row.isSection) return acc;
          if (row.readOnly && row.autoCalc) return acc + (supervisionTotals[row.autoCalc] || 0);
          const d = data[idx] || {};
          return acc + parseDollar(d.qty??0) * parseDollar(d.sub_up !== undefined ? d.sub_up : (row.sub_up??0));
        }, 0), ""],
      ];
      const ws = addSheet(name, rows);
      setColWidths(ws, [12, 42, 8, 8, 12, 14, 30]);
    });

    // ── Write file ────────────────────────────────────────────────────────────
    const fileName = `${project.jobNo}_${project.jobName.replace(/[^a-z0-9]/gi,"_")}_Estimate.xlsx`;
    XLSX.writeFile(wb, fileName);
  }


  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100%", overflow: "hidden" }}>
      {/* Workspace topbar */}
      <div style={{ padding: "8px 16px", borderBottom: `1px solid ${C.border}`, display: "flex", alignItems: "center", gap: 12, flexShrink: 0, backgroundColor: C.white }}>
        <button onClick={onBack} style={{ ...s.btn(), fontSize: 12, padding: "3px 10px" }}>← Projects</button>
        <div style={{ width: 1, height: 20, backgroundColor: C.border }} />
        <div style={{ fontSize: 13, fontWeight: 700, color: C.textPrimary }}>{project.jobName}</div>
        <div style={{ fontSize: 12, color: C.textMuted }}>{project.jobNo}</div>
        <div style={{ marginLeft: "auto", display: "flex", alignItems: "center", gap: 12 }}>
          <span style={{ fontSize: 11, color: saveStatus === "saving" ? C.textMuted : C.success, display: "flex", alignItems: "center", gap: 4 }}>
            {saveStatus === "saving"
              ? <span style={{ opacity: 0.6 }}>● Saving…</span>
              : <span>✓ Saved</span>}
          </span>
          <button
            onClick={handleExportExcel}
            style={{ ...s.btn("primary"), fontSize: 12, padding: "3px 10px" }}
          >
            ⬇ Export Excel
          </button>
          <button
            onClick={() => setShowDeleteConfirm(true)}
            style={{ ...s.btn(), fontSize: 12, padding: "3px 10px", color: C.danger, borderColor: C.danger }}
          >
            Delete Estimate
          </button>
        </div>
      </div>

      {/* Delete confirmation modal */}
      {showDeleteConfirm && (
        <div style={s.overlay} onClick={e => e.target === e.currentTarget && setShowDeleteConfirm(false)}>
          <div style={{ backgroundColor: C.white, borderRadius: 4, width: 420, boxShadow: "0 16px 48px rgba(0,0,0,0.25)", overflow: "hidden" }}>
            <div style={{ ...s.modalHeader, borderBottom: `1px solid ${C.border}` }}>
              <div style={s.modalTitle}>Delete Estimate</div>
              <button style={s.closeBtn} onClick={() => setShowDeleteConfirm(false)}>×</button>
            </div>
            <div style={s.modalBody}>
              <p style={{ fontSize: 13, color: C.textPrimary, margin: 0 }}>
                Are you sure you want to delete <strong>{project.jobName}</strong>? This will permanently remove all estimate data and cannot be undone.
              </p>
            </div>
            <div style={{ ...s.modalFooter, borderTop: `1px solid ${C.border}`, justifyContent: "flex-end", gap: 8 }}>
              <button style={s.btn()} onClick={() => setShowDeleteConfirm(false)}>Cancel</button>
              <button style={{ ...s.btn("primary"), backgroundColor: C.danger, borderColor: C.danger }} onClick={onDelete}>
                Delete Permanently
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Tabs */}
      <div style={{ display: "flex", borderBottom: `1px solid ${C.border}`, backgroundColor: C.white, flexShrink: 0, overflowX: "auto" }}>
        {TABS.map(tab => (
          <div
            key={tab.key}
            onClick={() => setActiveTab(tab.key)}
            style={{
              padding: "9px 18px", fontSize: 13, fontWeight: 500, cursor: "pointer", whiteSpace: "nowrap",
              borderBottom: activeTab === tab.key ? `2px solid ${C.accentLight}` : "2px solid transparent",
              color: activeTab === tab.key ? C.accentLight : C.textMuted,
              backgroundColor: C.white,
            }}
          >
            {tab.label}
          </div>
        ))}
        {activeTab === "supervision" && (
          <div style={{ marginLeft: "auto", display: "flex", alignItems: "center", paddingRight: 12 }}>
            <button onClick={() => setShowSettings(true)} style={{ ...s.btn(), fontSize: 12, padding: "3px 12px" }}>⚙ Assumptions</button>
          </div>
        )}
      </div>

      {/* Tab content */}
      <div style={{ flex: 1, overflow: "hidden", display: "flex", flexDirection: "column" }}>
        {activeTab === "recap"       && <RecapSheet project={project} recapData={recapData} onRecapChange={setRecapData} sheetTotals={sheetTotals} firmDefaults={firmDefaults} bidData={bidData} insSsRate={insSsRate} />}
        {activeTab === "bidsheet"    && <BidSheet bidData={bidData} onBidChange={setBidData} />}
        {activeTab === "supervision" && <SupervisionSheet supervisionData={supervisionData} onSupervisionChange={setSupervisionData} durationWeeks={durationWeeks} firmRates={firmRates} insSsRate={insSsRate} onInsSsRateChange={setInsSsRate} assumptions={assumptions} />}
        {activeTab === "gc"          && <DetailSheet title="General Conditions" code="2500" rows={SHEET_2500_ROWS} sheetData={sheet2500} onSheetChange={setSheet2500} autoCalcValues={{ supervision: supervisionTotals.supervision, perDiem: supervisionTotals.perDiem, travel: supervisionTotals.travel, fuel: supervisionTotals.fuel, safetyDir: supervisionTotals.safetyDir }} />}
        {activeTab === "equipment"   && <DetailSheet title="Equipment" code="2600" rows={SHEET_2600_ROWS} sheetData={sheet2600} onSheetChange={setSheet2600} autoCalcValues={{ fuel: supervisionTotals.fuel }} />}
        {activeTab === "insurance"   && <DetailSheet title="Insurance, Permits & Bonding" code="2700" rows={SHEET_2700_ROWS} sheetData={sheet2700} onSheetChange={setSheet2700} autoCalcValues={{}} />}
        {activeTab === "allowances"  && <DetailSheet title="Allowances / Testing & Inspection" code="2900" rows={SHEET_2900_ROWS} sheetData={sheet2900} onSheetChange={setSheet2900} autoCalcValues={{}} />}
        {activeTab === "sov"         && <SOVSheet project={project} recapData={recapData} bidData={bidData} sheetTotals={sheetTotals} firmDefaults={firmDefaults} insSsRate={insSsRate} />}
      </div>

      {showSettings && (
        <SettingsModal
          firmRates={firmRates}
          assumptions={assumptions}
          insSsRate={insSsRate}
          onSave={(rates, newAssump, newInsSs) => { onFirmRatesChange(rates); setAssumptions(newAssump); setInsSsRate(newInsSs); }}
          onClose={() => setShowSettings(false)}
        />
      )}
    </div>
  );
}

// ── localStorage helpers ─────────────────────────────────────────────────────
const getBidYear = (p) => p.dueDate?.slice(0,4) || p.createdAt?.slice(0,4) || "Unknown";

const LS_PROJECTS    = "mcg_projects_v1";
const LS_FIRM_RATES  = "mcg_firm_rates_v1";
const LS_FIRM_DEFS   = "mcg_firm_defaults_v1";
const lsGet = (key, fallback) => { try { const v = localStorage.getItem(key); return v ? JSON.parse(v) : fallback; } catch { return fallback; } };
const lsSet = (key, val)      => { try { localStorage.setItem(key, JSON.stringify(val)); } catch {} };
const lsEstKey = (id) => `mcg_estimate_${id}_v1`;

// ── Estimates Root ────────────────────────────────────────────────────────────
function Estimates() {
  const [projects,      setProjects]      = useState(() => {
    const saved = lsGet(LS_PROJECTS, []);
    return saved.map(p => {
      // Use cached value if present
      if (p.projectTotal) return p;
      // Otherwise compute from raw saved estimate data
      try {
        const est = lsGet(lsEstKey(p.id), {});
        if (!est || !est.recapData) return p;
        const { recapData={}, bidData={}, sheet2500={}, sheet2600={}, sheet2700={}, sheet2900={}, insSsRate } = est;
        const fR = parseDollar(recapData.__feeRate        ?? (FIRM_DEFAULTS.feeRate*100));
        const bR = parseDollar(recapData.__bondRate       ?? (FIRM_DEFAULTS.bondRate*100));
        const gR = parseDollar(recapData.__glRate         ?? (FIRM_DEFAULTS.glRate*100));
        const tR = parseDollar(recapData.__grtRate        ?? (FIRM_DEFAULTS.grtRate*100));
        const cA = parseDollar(recapData.__contingencyAmt ?? 0);
        const iR = parseDollar(insSsRate ?? (FIRM_DEFAULTS.insSsRate*100));
        // Bid amounts by recap code
        const bidByCode = {};
        BIDSHEET_TRADES.forEach((trade, idx) => {
          const d = (bidData||{})[idx] || {};
          const amt = parseDollar(d.amount ?? 0); if (!amt) return;
          const type = d.type ?? trade.type;
          const k5 = trade.code.slice(0,5), k2 = trade.code.slice(0,2);
          const rk = RECAP_TRADE_ROWS.find(r=>r.code===k5)?.code || RECAP_TRADE_ROWS.find(r=>r.code.slice(0,2)===k2)?.code || null;
          if (!rk) return;
          if (!bidByCode[rk]) bidByCode[rk] = { material:0, sub:0 };
          if (type==="M") bidByCode[rk].material += amt;
          else if (type==="S") bidByCode[rk].sub += amt;
        });
        // Trade totals
        let tradeSum=0, laborSum=0;
        RECAP_TRADE_ROWS.forEach(row => {
          const d = recapData[row.code]||{}, bid = bidByCode[row.code]||{};
          const l = parseDollar(d.labor!==undefined&&d.labor!==""?d.labor:0);
          const m = parseDollar(d.material!==undefined&&d.material!==""?d.material:(bid.material??0));
          const sb= parseDollar(d.sub!==undefined&&d.sub!==""?d.sub:(bid.sub??0));
          tradeSum += l+m+sb; laborSum += l;
        });
        // GC sheet totals (simplified — qty × sub_up)
        const calcSheet = (rows, data) => rows.reduce((acc, row, idx) => {
          if (row.isSection || (row.readOnly && row.autoCalc)) return acc;
          const d = data[idx]||{};
          return acc + parseDollar(d.qty??0)*parseDollar(d.sub_up!==undefined?d.sub_up:(row.sub_up??0));
        }, 0);
        const gcTotal = calcSheet(SHEET_2500_ROWS,sheet2500)+calcSheet(SHEET_2600_ROWS,sheet2600)+calcSheet(SHEET_2700_ROWS,sheet2700)+calcSheet(SHEET_2900_ROWS,sheet2900);
        const sub = tradeSum+gcTotal;
        const ins = laborSum*(iR/100);
        const s2  = sub+ins;
        const fee = s2*(fR/100), bond=(s2+fee)*(bR/100), gl=(s2+fee)*(gR/100), grt=(s2+fee)*(tR/100);
        const total = s2+fee+bond+gl+grt+cA;
        return total > 0 ? { ...p, projectTotal: total } : p;
      } catch { return p; }
    });
  });
  const [activeProject,  setActiveProject]  = useState(null);
  const [showNewModal,   setShowNewModal]   = useState(false);
  const [editingProject, setEditingProject] = useState(null);
  const [firmRates,     setFirmRates]     = useState(() => lsGet(LS_FIRM_RATES,
    Object.fromEntries(SUPERVISION_ROLES.map(r => [r.key, r.defaultRate]))
  ));
  const [firmDefaults, setFirmDefaults]   = useState(() => lsGet(LS_FIRM_DEFS, { ...FIRM_DEFAULTS }));

  // Persist projects list whenever it changes
  useEffect(() => { lsSet(LS_PROJECTS, projects); }, [projects]);
  useEffect(() => { lsSet(LS_FIRM_RATES, firmRates); }, [firmRates]);
  useEffect(() => { lsSet(LS_FIRM_DEFS, firmDefaults); }, [firmDefaults]);

  function handleCreate(project) {
    setProjects(prev => [project, ...prev]);
    setShowNewModal(false);
    setActiveProject(project);
  }

  function handleDeleteProject(id) {
    setProjects(prev => prev.filter(p => p.id !== id));
    try { localStorage.removeItem(lsEstKey(id)); } catch {}
    setActiveProject(null);
  }

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100%", overflow: "hidden" }}>
      {activeProject ? (
        <EstimateWorkspace
          key={activeProject.id}
          project={activeProject}
          onBack={() => setActiveProject(null)}
          onDelete={() => handleDeleteProject(activeProject.id)}
          onProjectUpdate={updates => setProjects(prev => prev.map(p => p.id === activeProject.id ? { ...p, ...updates } : p))}
          firmRates={firmRates}
          firmDefaults={firmDefaults}
          onFirmRatesChange={setFirmRates}
          onFirmDefaultsChange={setFirmDefaults}
        />
      ) : (
        <ProjectList
          projects={projects}
          onSelect={setActiveProject}
          onNew={() => setShowNewModal(true)}
          onStatusChange={(id, status) => setProjects(prev => prev.map(p => p.id === id ? { ...p, status } : p))}
          onEdit={p => setEditingProject(p)}
        />
      )}
      {showNewModal && (
        <NewProjectModal
          onClose={() => setShowNewModal(false)}
          onCreate={handleCreate}
        />
      )}
      {editingProject && (
        <EditProjectModal
          project={editingProject}
          onClose={() => setEditingProject(null)}
          onSave={updated => {
            setProjects(prev => prev.map(p => p.id === updated.id ? { ...p, ...updated } : p));
            setEditingProject(null);
          }}
        />
      )}
    </div>
  );
}


// ─── Bid Comparison ───────────────────────────────────────────────────────────

// Compute full project total + all trade/GC line items from raw saved estimate
function computeEstimateSnapshot(p) {
  try {
    const est = lsGet(lsEstKey(p.id), {});
    if (!est) return null;
    const { recapData={}, bidData={}, sheet2500={}, sheet2600={}, sheet2700={}, sheet2900={}, insSsRate } = est;

    const fR = parseDollar(recapData.__feeRate        ?? (FIRM_DEFAULTS.feeRate*100));
    const bR = parseDollar(recapData.__bondRate       ?? (FIRM_DEFAULTS.bondRate*100));
    const gR = parseDollar(recapData.__glRate         ?? (FIRM_DEFAULTS.glRate*100));
    const tR = parseDollar(recapData.__grtRate        ?? (FIRM_DEFAULTS.grtRate*100));
    const cA = parseDollar(recapData.__contingencyAmt ?? 0);
    const iR = parseDollar(insSsRate ?? (FIRM_DEFAULTS.insSsRate*100));

    // Bid amounts by recap code
    const bidByCode = {};
    BIDSHEET_TRADES.forEach((trade, idx) => {
      const d = (bidData||{})[idx] || {};
      const amt = parseDollar(d.amount ?? 0); if (!amt) return;
      const type = d.type ?? trade.type;
      const k5 = trade.code.slice(0,5), k2 = trade.code.slice(0,2);
      const rk = RECAP_TRADE_ROWS.find(r=>r.code===k5)?.code
              || RECAP_TRADE_ROWS.find(r=>r.code.slice(0,2)===k2)?.code || null;
      if (!rk) return;
      if (!bidByCode[rk]) bidByCode[rk] = { material:0, sub:0 };
      if (type==="M") bidByCode[rk].material += amt;
      else if (type==="S") bidByCode[rk].sub += amt;
    });

    // Per-trade totals
    const tradeData = {};
    let tradeSum=0, laborSum=0;
    RECAP_TRADE_ROWS.forEach(row => {
      const d = recapData[row.code]||{}, bid = bidByCode[row.code]||{};
      const l = parseDollar(d.labor!==undefined&&d.labor!==""?d.labor:0);
      const m = parseDollar(d.material!==undefined&&d.material!==""?d.material:(bid.material??0));
      const sb= parseDollar(d.sub!==undefined&&d.sub!==""?d.sub:(bid.sub??0));
      const total = l+m+sb;
      tradeData[row.code] = { labor:l, material:m, sub:sb, total };
      tradeSum += total; laborSum += l;
    });

    // GC sheets
    const calcSheet = (rows, data) => rows.reduce((acc, row, idx) => {
      if (row.isSection || (row.readOnly && row.autoCalc)) return acc;
      const d = data[idx]||{};
      return acc + parseDollar(d.qty??0)*parseDollar(d.sub_up!==undefined?d.sub_up:(row.sub_up??0));
    }, 0);
    const gc2500 = calcSheet(SHEET_2500_ROWS,sheet2500);
    const gc2600 = calcSheet(SHEET_2600_ROWS,sheet2600);
    const gc2700 = calcSheet(SHEET_2700_ROWS,sheet2700);
    const gc2900 = calcSheet(SHEET_2900_ROWS,sheet2900);
    const gcTotal = gc2500+gc2600+gc2700+gc2900;

    const sub    = tradeSum+gcTotal;
    const ins    = laborSum*(iR/100);
    const s2     = sub+ins;
    const fee    = s2*(fR/100);
    const bond   = (s2+fee)*(bR/100);
    const gl     = (s2+fee)*(gR/100);
    const grt    = (s2+fee)*(tR/100);
    const total  = s2+fee+bond+gl+grt+cA;

    return {
      tradeData,
      gc: { "2500":gc2500, "2600":gc2600, "2700":gc2700, "2900":gc2900, total:gcTotal },
      subtotal: s2,
      fee, bond, gl, grt, contingency:cA,
      projectTotal: total,
    };
  } catch(e) { return null; }
}

function BidComparison() {
  // Load all projects from localStorage
  const allProjects = useMemo(() => lsGet(LS_PROJECTS, []), []);

  // Filters
  const [filterType,    setFilterType]    = useState("All");
  const [filterStatus,  setFilterStatus]  = useState("All");
  const [filterMarket,  setFilterMarket]  = useState("All");
  const [selectedIds,   setSelectedIds]   = useState([]);
  const [activeRow,     setActiveRow]     = useState(null); // for drill-down highlight
  const [metric,        setMetric]        = useState("psf"); // "psf" | "total" | "pct"

  // Derived filter options
  const typeOptions    = useMemo(() => ["All", ...new Set(allProjects.map(p => p.projectType).filter(Boolean))], [allProjects]);
  const statusOptions  = useMemo(() => ["All", ...BID_STATUSES], []);
  const marketOptions  = useMemo(() => ["All", ...new Set(allProjects.map(p => p.location).filter(Boolean))], [allProjects]);

  // Filtered project pool
  const filteredProjects = useMemo(() => allProjects.filter(p => {
    if (filterType   !== "All" && p.projectType !== filterType)  return false;
    if (filterStatus !== "All" && p.status      !== filterStatus) return false;
    if (filterMarket !== "All" && p.location    !== filterMarket) return false;
    return true;
  }), [allProjects, filterType, filterStatus, filterMarket]);

  // Auto-select first two when filter changes
  useEffect(() => {
    setSelectedIds(filteredProjects.slice(0, 2).map(p => p.id));
  }, [filteredProjects.map(p=>p.id).join(",")]);

  // Selected projects with computed snapshots
  const selectedProjects = useMemo(() =>
    selectedIds.map(id => {
      const p = allProjects.find(p => p.id === id);
      if (!p) return null;
      const snap = computeEstimateSnapshot(p);
      return { ...p, snap };
    }).filter(Boolean),
  [selectedIds, allProjects]);

  function toggleProject(id) {
    setSelectedIds(prev =>
      prev.includes(id) ? prev.filter(x => x !== id) : [...prev, id]
    );
  }

  // Compute cell value for a trade row
  function getVal(proj, code) {
    const snap = proj.snap;
    if (!snap) return null;
    const td = snap.tradeData[code];
    if (!td || td.total === 0) return null;
    const sqft = parseDollar(proj.sqft || 0);
    if (metric === "psf")   return sqft > 0 ? td.total / sqft : null;
    if (metric === "total") return td.total;
    if (metric === "pct")   return snap.projectTotal > 0 ? (td.total / snap.projectTotal) * 100 : null;
    return null;
  }

  function getGCVal(proj, key) {
    const snap = proj.snap;
    if (!snap) return null;
    const val = key === "total" ? snap.gc.total : snap.gc[key];
    if (!val) return null;
    const sqft = parseDollar(proj.sqft || 0);
    if (metric === "psf")   return sqft > 0 ? val / sqft : null;
    if (metric === "total") return val;
    if (metric === "pct")   return snap.projectTotal > 0 ? (val / snap.projectTotal) * 100 : null;
    return null;
  }

  function getTotalVal(proj) {
    const snap = proj.snap;
    if (!snap || !snap.projectTotal) return null;
    const sqft = parseDollar(proj.sqft || 0);
    if (metric === "psf")   return sqft > 0 ? snap.projectTotal / sqft : null;
    if (metric === "total") return snap.projectTotal;
    if (metric === "pct")   return 100;
    return null;
  }

  // Outlier detection — flag cells > 20% from group avg
  function getAvg(code, isGC=false, gcKey=null) {
    const vals = selectedProjects.map(p => isGC ? getGCVal(p, gcKey||code) : getVal(p, code)).filter(v => v != null && v > 0);
    if (vals.length < 2) return null;
    return vals.reduce((a,v)=>a+v,0) / vals.length;
  }

  function outlierStyle(val, avg) {
    if (val == null || avg == null || avg === 0) return {};
    const diff = (val - avg) / avg;
    if (diff >  0.20) return { backgroundColor: "#FEF3C7", color: "#92400E" };
    if (diff < -0.20) return { backgroundColor: "#DCFCE7", color: "#166534" };
    return {};
  }

  function fmtVal(val) {
    if (val == null) return "—";
    if (metric === "psf")   return "$" + val.toFixed(2);
    if (metric === "total") return "$" + Number(val).toLocaleString("en-US", { minimumFractionDigits: 0, maximumFractionDigits: 0 });
    if (metric === "pct")   return val.toFixed(1) + "%";
    return "—";
  }

  const thBase = {
    padding: "8px 14px", fontSize: 11, fontWeight: 700, color: C.textMuted,
    textTransform: "uppercase", letterSpacing: "0.06em",
    borderBottom: `2px solid ${C.border}`, textAlign: "right",
    backgroundColor: C.contentBg, whiteSpace: "nowrap",
    position: "sticky", top: 0, zIndex: 2,
  };
  const tdBase = { padding: "6px 14px", fontSize: 12, borderBottom: `1px solid ${C.border}`, textAlign: "right", fontFamily: "monospace" };
  const labelTd = { padding: "6px 14px", fontSize: 12, borderBottom: `1px solid ${C.border}`, color: C.textPrimary, position: "sticky", left: 0, backgroundColor: C.white, zIndex: 1 };
  const codeTd  = { padding: "6px 14px", fontSize: 11, borderBottom: `1px solid ${C.border}`, color: C.textMuted, fontFamily: "monospace", position: "sticky", left: 0, backgroundColor: C.white, zIndex: 1 };

  const sectionHdr = (label) => (
    <tr>
      <td colSpan={selectedProjects.length + 1} style={{
        padding: "8px 14px 4px", fontSize: 10, fontWeight: 800, color: C.textMuted,
        textTransform: "uppercase", letterSpacing: "0.1em",
        backgroundColor: C.contentBg, borderBottom: `1px solid ${C.border}`,
      }}>{label}</td>
    </tr>
  );

  return (
    <div style={{ display:"flex", flexDirection:"column", height:"100%", overflow:"hidden" }}>

      {/* Header */}
      <div style={{ padding:"12px 24px 10px", borderBottom:`1px solid ${C.border}`, flexShrink:0, backgroundColor:C.white }}>
        <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:12 }}>
          <div>
            <div style={{ fontSize:15, fontWeight:700, color:C.textPrimary }}>Bid Comparison</div>
            <div style={{ fontSize:12, color:C.textMuted, marginTop:2 }}>
              Compare estimates side by side — {selectedProjects.length} selected
            </div>
          </div>
          {/* Metric toggle */}
          <div style={{ display:"flex", gap:2, backgroundColor:C.contentBg, borderRadius:4, padding:3, border:`1px solid ${C.border}` }}>
            {[["psf","$/SF"],["total","Total $"],["pct","% of Total"]].map(([key,label]) => (
              <button key={key} onClick={() => setMetric(key)} style={{
                padding:"4px 14px", fontSize:12, fontWeight:600, borderRadius:3, border:"none", cursor:"pointer",
                backgroundColor: metric===key ? C.accentLight : "transparent",
                color: metric===key ? C.white : C.textMuted,
              }}>{label}</button>
            ))}
          </div>
        </div>

        {/* Filters */}
        <div style={{ display:"flex", gap:12, alignItems:"flex-end", flexWrap:"wrap" }}>
          {[
            ["Type",   filterType,   setFilterType,   typeOptions],
            ["Status", filterStatus, setFilterStatus, statusOptions],
            ["Location", filterMarket, setFilterMarket, marketOptions],
          ].map(([label, val, setter, opts]) => (
            <div key={label} style={{ display:"flex", flexDirection:"column", gap:3 }}>
              <label style={{ fontSize:10, fontWeight:700, color:C.textMuted, textTransform:"uppercase", letterSpacing:"0.07em" }}>{label}</label>
              <select value={val} onChange={e => setter(e.target.value)} style={{ ...s.fieldSelect, width:140, fontSize:12 }}>
                {opts.map(o => <option key={o} value={o}>{o}</option>)}
              </select>
            </div>
          ))}
          <div style={{ width:1, height:32, backgroundColor:C.border, margin:"0 4px" }} />
          <div style={{ display:"flex", flexDirection:"column", gap:3 }}>
            <label style={{ fontSize:10, fontWeight:700, color:C.textMuted, textTransform:"uppercase", letterSpacing:"0.07em" }}>Projects in comparison</label>
            <div style={{ display:"flex", gap:6, flexWrap:"wrap" }}>
              {filteredProjects.map(p => {
                const sel = selectedIds.includes(p.id);
                return (
                  <button key={p.id} onClick={() => toggleProject(p.id)} style={{
                    padding:"3px 10px", fontSize:11, fontWeight:600, borderRadius:3, cursor:"pointer",
                    border:`1px solid ${sel ? C.accentLight : C.border}`,
                    backgroundColor: sel ? "#EFF6FF" : C.white,
                    color: sel ? C.accentLight : C.textMuted,
                  }}>
                    {p.jobNo} — {p.jobName}
                  </button>
                );
              })}
              {filteredProjects.length === 0 && <span style={{ fontSize:12, color:C.textMuted, fontStyle:"italic" }}>No projects match filters</span>}
            </div>
          </div>
        </div>
      </div>

      {/* Comparison table */}
      {selectedProjects.length === 0 ? (
        <div style={{ flex:1, display:"flex", alignItems:"center", justifyContent:"center", color:C.textMuted, flexDirection:"column", gap:10 }}>
          <div style={{ fontSize:32 }}>⇄</div>
          <div style={{ fontWeight:600, fontSize:14, color:C.textPrimary }}>Select projects to compare</div>
          <div style={{ fontSize:13 }}>Choose two or more projects from the filters above.</div>
        </div>
      ) : (
        <div style={{ flex:1, overflow:"auto" }}>
          <table style={{ borderCollapse:"collapse", minWidth:600 }}>
            <thead>
              <tr>
                {/* Sticky label header */}
                <th style={{ ...thBase, textAlign:"left", position:"sticky", left:0, zIndex:3, width:320 }}>
                  Trade / Line Item
                </th>
                {selectedProjects.map(p => (
                  <th key={p.id} style={{ ...thBase, minWidth:160 }}>
                    <div style={{ fontSize:11, fontWeight:700, color:C.textPrimary }}>{p.jobName}</div>
                    <div style={{ fontSize:10, color:C.textMuted, fontWeight:400 }}>{p.jobNo} · {p.sqft ? Number(p.sqft).toLocaleString()+" SF" : "—"}</div>
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {/* Project meta row */}
              <tr style={{ backgroundColor:"#F8FAFF" }}>
                <td style={{ ...labelTd, fontWeight:700, backgroundColor:"#F8FAFF" }}>Project Total</td>
                {selectedProjects.map(p => {
                  const v = getTotalVal(p);
                  return (
                    <td key={p.id} style={{ ...tdBase, fontWeight:700, color: v ? C.accentLight : C.textLight }}>
                      {fmtVal(v)}
                    </td>
                  );
                })}
              </tr>

              {/* ── Trade rows ── */}
              {sectionHdr("Trade Work")}
              {RECAP_TRADE_ROWS.map((row, ri) => {
                const avg = getAvg(row.code);
                const rowVals = selectedProjects.map(p => getVal(p, row.code));
                const hasAny  = rowVals.some(v => v != null);
                if (!hasAny) return null;
                const isActive = activeRow === row.code;
                return (
                  <tr key={row.code}
                    style={{ backgroundColor: isActive ? "#EFF6FF" : ri%2===0 ? C.white : C.tableStripe, cursor:"pointer" }}
                    onClick={() => setActiveRow(isActive ? null : row.code)}
                  >
                    <td style={{ ...labelTd, backgroundColor: isActive ? "#EFF6FF" : ri%2===0 ? C.white : C.tableStripe }}>
                      <div style={{ display:"flex", alignItems:"baseline", gap:8 }}>
                        <span style={{ fontSize:10, fontFamily:"monospace", color:C.textMuted }}>{row.code}</span>
                        <span style={{ fontSize:12, color:C.textPrimary }}>{row.label}</span>
                      </div>
                    </td>
                    {selectedProjects.map(p => {
                      const v = getVal(p, row.code);
                      return (
                        <td key={p.id} style={{ ...tdBase, ...outlierStyle(v, avg), color: v ? C.textPrimary : C.textLight }}>
                          {fmtVal(v)}
                        </td>
                      );
                    })}
                  </tr>
                );
              })}

              {/* ── GC Sheets ── */}
              {sectionHdr("General Conditions")}
              {[
                { key:"2500", label:"General Conditions" },
                { key:"2600", label:"Equipment" },
                { key:"2700", label:"Insurance & Permits" },
                { key:"2900", label:"Allowances / Testing" },
                { key:"total", label:"Total GC" },
              ].map((gc, gi) => {
                const avg = getAvg(null, true, gc.key);
                const rowVals = selectedProjects.map(p => getGCVal(p, gc.key));
                const hasAny = rowVals.some(v => v != null);
                if (!hasAny) return null;
                return (
                  <tr key={gc.key} style={{ backgroundColor: gc.key==="total" ? "#F0F4FF" : gi%2===0 ? C.white : C.tableStripe }}>
                    <td style={{ ...labelTd, fontWeight: gc.key==="total" ? 700 : 400, backgroundColor: gc.key==="total" ? "#F0F4FF" : gi%2===0 ? C.white : C.tableStripe }}>
                      <div style={{ display:"flex", alignItems:"baseline", gap:8 }}>
                        {gc.key !== "total" && <span style={{ fontSize:10, fontFamily:"monospace", color:C.textMuted }}>{gc.key}</span>}
                        <span style={{ fontSize:12, color:C.textPrimary }}>{gc.label}</span>
                      </div>
                    </td>
                    {selectedProjects.map(p => {
                      const v = getGCVal(p, gc.key);
                      return (
                        <td key={p.id} style={{ ...tdBase, fontWeight: gc.key==="total"?700:400, ...outlierStyle(v, avg), color: v ? C.textPrimary : C.textLight }}>
                          {fmtVal(v)}
                        </td>
                      );
                    })}
                  </tr>
                );
              })}

              {/* ── Fee / Bond / GL / GRT ── */}
              {sectionHdr("Markup & Fees")}
              {[
                { key:"fee",         label:"GC Fee" },
                { key:"bond",        label:"Performance & Payment Bond" },
                { key:"gl",          label:"General Liability Insurance" },
                { key:"grt",         label:"Gross Receipts Tax" },
                { key:"contingency", label:"Contingency" },
              ].map((mk, mi) => {
                const rowVals = selectedProjects.map(p => {
                  const snap = p.snap; if (!snap) return null;
                  const v = snap[mk.key]; if (!v) return null;
                  const sqft = parseDollar(p.sqft||0);
                  if (metric==="psf")   return sqft>0 ? v/sqft : null;
                  if (metric==="total") return v;
                  if (metric==="pct")   return snap.projectTotal>0 ? (v/snap.projectTotal)*100 : null;
                  return null;
                });
                const hasAny = rowVals.some(v => v != null);
                if (!hasAny) return null;
                return (
                  <tr key={mk.key} style={{ backgroundColor: mi%2===0 ? C.white : C.tableStripe }}>
                    <td style={{ ...labelTd, color:C.textMuted, fontStyle:"italic", backgroundColor: mi%2===0 ? C.white : C.tableStripe }}>
                      <span style={{ fontSize:12 }}>{mk.label}</span>
                    </td>
                    {rowVals.map((v,i) => (
                      <td key={i} style={{ ...tdBase, color: v ? C.textMuted : C.textLight }}>{fmtVal(v)}</td>
                    ))}
                  </tr>
                );
              })}

              {/* ── Project Total footer ── */}
              <tr style={{ backgroundColor: C.accent }}>
                <td style={{ ...labelTd, fontWeight:800, fontSize:13, color:C.white, backgroundColor:C.accent, borderTop:`2px solid ${C.border}` }}>
                  TOTAL PROPOSED CONTRACT SUM
                </td>
                {selectedProjects.map(p => {
                  const v = getTotalVal(p);
                  return (
                    <td key={p.id} style={{ ...tdBase, fontWeight:800, fontSize:13, color:C.white, borderTop:`2px solid ${C.border}` }}>
                      {fmtVal(v)}
                    </td>
                  );
                })}
              </tr>
            </tbody>
          </table>

          {/* Legend */}
          <div style={{ padding:"12px 20px", display:"flex", gap:20, fontSize:11, color:C.textMuted, borderTop:`1px solid ${C.border}` }}>
            <span style={{ display:"flex", alignItems:"center", gap:6 }}>
              <span style={{ width:12, height:12, backgroundColor:"#FEF3C7", border:"1px solid #D97706", borderRadius:2, display:"inline-block" }} />
              20%+ above group average
            </span>
            <span style={{ display:"flex", alignItems:"center", gap:6 }}>
              <span style={{ width:12, height:12, backgroundColor:"#DCFCE7", border:"1px solid #16A34A", borderRadius:2, display:"inline-block" }} />
              20%+ below group average
            </span>
            <span style={{ color:C.textLight, marginLeft:"auto" }}>Click a trade row to highlight it for reference</span>
          </div>
        </div>
      )}
    </div>
  );
}


// ─── Dashboard ────────────────────────────────────────────────────────────────
function Dashboard() {
  const allProjects = useMemo(() => lsGet(LS_PROJECTS, []), []);
  const years = useMemo(() => {
    // Bid year: prefer due date year, fall back to createdAt year
    const ys = [...new Set(allProjects.map(p => getBidYear(p)).filter(y => y !== 'Unknown'))].sort().reverse();
    return ["All", ...ys];
  }, [allProjects]);

  const [yearFilter, setYearFilter] = useState("All");

  const projects = useMemo(() => {
    if (yearFilter === "All") return allProjects;
    return allProjects.filter(p => getBidYear(p) === yearFilter);
  }, [allProjects, yearFilter]);

  const today = new Date();

  // ── Core metrics ────────────────────────────────────────────────────────────
  const metrics = useMemo(() => {
    const decided = projects.filter(p => p.status === "Won" || p.status === "Lost");
    const won     = projects.filter(p => p.status === "Won");
    const open    = projects.filter(p => p.status === "Bidding" || p.status === "Pending");
    const hitRate = decided.length > 0 ? won.length / decided.length : 0;
    const revenueWon  = won.reduce((a, p) => a + (p.projectTotal || 0), 0);
    const projFeeWon  = won.reduce((a, p) => a + (parseDollar(p.projectedFee) || 0), 0);
    const revPursued  = decided.reduce((a, p) => a + (p.projectTotal || 0), 0);
    return { hitRate, jobsWon: won.length, revenueWon, projFeeWon, openBids: open.length, totalBids: projects.length, revPursued };
  }, [projects]);

  // ── Upcoming bids (Bidding/Pending with due date) ────────────────────────
  const upcomingBids = useMemo(() => {
    return allProjects
      .filter(p => (p.status === "Bidding" || p.status === "Pending") && p.dueDate)
      .map(p => {
        const due     = new Date(p.dueDate);
        const daysOut = Math.ceil((due - today) / (1000*60*60*24));
        return { ...p, daysOut };
      })
      .filter(p => p.daysOut >= 0 && p.daysOut <= 60)
      .sort((a,b) => a.daysOut - b.daysOut);
  }, [allProjects, today]);

  // ── Industry breakdown ───────────────────────────────────────────────────
  const byIndustry = useMemo(() => {
    const map = {};
    projects.forEach(p => {
      const ind = p.industry || "Other";
      if (!map[ind]) map[ind] = { won:0, lost:0, revWon:0, feeWon:0 };
      if (p.status === "Won")  { map[ind].won++;  map[ind].revWon += p.projectTotal||0; map[ind].feeWon += parseDollar(p.projectedFee)||0; }
      if (p.status === "Lost") { map[ind].lost++; }
    });
    return Object.entries(map)
      .map(([ind, d]) => ({ ind, ...d, hitRate: (d.won+d.lost)>0 ? d.won/(d.won+d.lost) : null }))
      .filter(r => r.won+r.lost > 0)
      .sort((a,b) => b.revWon - a.revWon);
  }, [projects]);

  // ── Bid type breakdown ───────────────────────────────────────────────────
  const byBidType = useMemo(() => {
    const map = {};
    projects.filter(p => p.status==="Won"||p.status==="Lost").forEach(p => {
      const t = p.bidType || "Hard Bid";
      if (!map[t]) map[t] = { won:0, lost:0, revWon:0 };
      if (p.status==="Won")  { map[t].won++;  map[t].revWon += p.projectTotal||0; }
      if (p.status==="Lost") { map[t].lost++; }
    });
    return Object.entries(map).map(([type, d]) => ({ type, ...d })).sort((a,b)=>b.revWon-a.revWon);
  }, [projects]);

  // ── Status breakdown ─────────────────────────────────────────────────────
  const byStatus = useMemo(() => {
    const map = {};
    projects.forEach(p => {
      if (!map[p.status]) map[p.status] = { count:0, rev:0, fee:0 };
      map[p.status].count++;
      map[p.status].rev += p.projectTotal||0;
      map[p.status].fee += parseDollar(p.projectedFee)||0;
    });
    return BID_STATUSES.map(st => ({ status:st, ...(map[st]||{count:0,rev:0,fee:0}) }));
  }, [projects]);

  // ── Top clients ──────────────────────────────────────────────────────────
  const topClients = useMemo(() => {
    const map = {};
    projects.filter(p=>p.status==="Won"&&p.client).forEach(p => {
      if (!map[p.client]) map[p.client] = { won:0, rev:0, fee:0 };
      map[p.client].won++;
      map[p.client].rev += p.projectTotal||0;
      map[p.client].fee += parseDollar(p.projectedFee)||0;
    });
    return Object.entries(map).map(([client,d])=>({client,...d})).sort((a,b)=>b.rev-a.rev).slice(0,8);
  }, [projects]);

  // ── YoY summary ─────────────────────────────────────────────────────────
  const yoy = useMemo(() => {
    const map = {};
    allProjects.forEach(p => {
      const yr = getBidYear(p);
      if (!map[yr]) map[yr] = { total:0, won:0, lost:0, revWon:0 };
      map[yr].total++;
      if (p.status==="Won")  { map[yr].won++;  map[yr].revWon += p.projectTotal||0; }
      if (p.status==="Lost") { map[yr].lost++; }
    });
    return Object.entries(map)
      .map(([yr,d]) => ({ yr, ...d, hitRate: (d.won+d.lost)>0 ? d.won/(d.won+d.lost) : null }))
      .sort((a,b)=>b.yr.localeCompare(a.yr));
  }, [allProjects]);

  // ── Styles ───────────────────────────────────────────────────────────────
  const fmtM = (v) => v >= 1000000 ? `$${(v/1000000).toFixed(1)}M` : v >= 1000 ? `$${(v/1000).toFixed(0)}K` : `$${v.toFixed(0)}`;
  const fmtPct = (v) => v == null ? "—" : (v*100).toFixed(1) + "%";
  const fmtRev = (v) => "$" + Number(v||0).toLocaleString("en-US", {minimumFractionDigits:0,maximumFractionDigits:0});

  const tH = { padding:"7px 14px", fontSize:11, fontWeight:700, color:C.textMuted, textTransform:"uppercase", letterSpacing:"0.06em", borderBottom:`2px solid ${C.border}`, textAlign:"left", backgroundColor:C.contentBg, whiteSpace:"nowrap" };
  const tR = { padding:"7px 14px", fontSize:12, borderBottom:`1px solid ${C.border}` };
  const tRr = { ...tR, textAlign:"right", fontFamily:"monospace" };

  const urgencyColor = (days) => {
    if (days <= 7)  return { bg:"#FEE2E2", text:"#991B1B", badge:"#DC2626" };
    if (days <= 14) return { bg:"#FEF3C7", text:"#92400E", badge:"#D97706" };
    return { bg:"#DCFCE7", text:"#166534", badge:"#16A34A" };
  };

  return (
    <div style={{ display:"flex", flexDirection:"column", height:"100%", overflow:"hidden" }}>

      {/* Header */}
      <div style={{ padding:"12px 24px 10px", borderBottom:`1px solid ${C.border}`, flexShrink:0, backgroundColor:C.white, display:"flex", alignItems:"center", justifyContent:"space-between" }}>
        <div>
          <div style={{ fontSize:15, fontWeight:700, color:C.textPrimary }}>Bid Tracker Dashboard</div>
          <div style={{ fontSize:12, color:C.textMuted, marginTop:2 }}>Meridian Construction Group</div>
        </div>
        {/* Year filter */}
        <div style={{ display:"flex", alignItems:"center", gap:8 }}>
          <span style={{ fontSize:12, color:C.textMuted }}>Year</span>
          <div style={{ display:"flex", gap:2, backgroundColor:C.contentBg, borderRadius:4, padding:3, border:`1px solid ${C.border}` }}>
            {years.map(yr => (
              <button key={yr} onClick={() => setYearFilter(yr)} style={{
                padding:"3px 12px", fontSize:12, fontWeight:600, borderRadius:3, border:"none", cursor:"pointer",
                backgroundColor: yearFilter===yr ? C.accentLight : "transparent",
                color: yearFilter===yr ? C.white : C.textMuted,
              }}>{yr}</button>
            ))}
          </div>
        </div>
      </div>

      <div style={{ flex:1, overflowY:"auto", padding:"20px 24px", display:"flex", flexDirection:"column", gap:20 }}>

        {/* ── KPI Strip ── */}
        <div style={{ display:"grid", gridTemplateColumns:"repeat(6,1fr)", gap:12 }}>
          {[
            { label:"Hit Rate",      value: fmtPct(metrics.hitRate),                   sub: `${metrics.jobsWon} won / ${metrics.jobsWon + byStatus.find(s=>s.status==="Lost")?.count||0} decided`, accent: true },
            { label:"Jobs Won",      value: metrics.jobsWon,                            sub: `of ${metrics.totalBids} total bids` },
            { label:"Revenue Won",   value: fmtM(metrics.revenueWon),                  sub: fmtRev(metrics.revenueWon) },
            { label:"Projected Fee", value: fmtM(metrics.projFeeWon),                  sub: metrics.revenueWon>0 ? fmtPct(metrics.projFeeWon/metrics.revenueWon)+" margin" : "—" },
            { label:"Open Bids",     value: metrics.openBids,                           sub: "Bidding + Pending" },
            { label:"Total Bids",    value: metrics.totalBids,                          sub: yearFilter==="All" ? "All time" : yearFilter },
          ].map(({ label, value, sub, accent }) => (
            <div key={label} style={{
              backgroundColor: accent ? C.accentLight : C.white,
              border: `1px solid ${accent ? C.accentLight : C.border}`,
              borderRadius:4, padding:"14px 16px",
            }}>
              <div style={{ fontSize:11, fontWeight:700, color: accent ? "rgba(255,255,255,0.8)" : C.textMuted, textTransform:"uppercase", letterSpacing:"0.07em", marginBottom:6 }}>{label}</div>
              <div style={{ fontSize:24, fontWeight:800, color: accent ? C.white : C.textPrimary, fontFamily:"monospace", lineHeight:1 }}>{value}</div>
              <div style={{ fontSize:11, color: accent ? "rgba(255,255,255,0.7)" : C.textMuted, marginTop:4 }}>{sub}</div>
            </div>
          ))}
        </div>

        {/* ── Upcoming Bids ── */}
        {upcomingBids.length > 0 && (
          <div style={{ backgroundColor:C.white, border:`1px solid ${C.border}`, borderRadius:4 }}>
            <div style={{ padding:"12px 16px 8px", borderBottom:`1px solid ${C.border}`, display:"flex", alignItems:"center", justifyContent:"space-between" }}>
              <div style={{ fontSize:13, fontWeight:700, color:C.textPrimary }}>📌 Upcoming Bids — Next 60 Days</div>
              <div style={{ fontSize:11, color:C.textMuted, display:"flex", gap:16 }}>
                {[["≤7 days","#DC2626"],["8–14 days","#D97706"],["15–60 days","#16A34A"]].map(([l,c])=>(
                  <span key={l} style={{ display:"flex", alignItems:"center", gap:4 }}>
                    <span style={{ width:8, height:8, borderRadius:"50%", backgroundColor:c, display:"inline-block" }}/>
                    {l}
                  </span>
                ))}
              </div>
            </div>
            <table style={{ width:"100%", borderCollapse:"collapse" }}>
              <thead>
                <tr>
                  {["Bid No.","Project","Client","Location","Bid Type","Due Date","Days Out","Status"].map(h=>(
                    <th key={h} style={tH}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {upcomingBids.map((p,i) => {
                  const uc = urgencyColor(p.daysOut);
                  return (
                    <tr key={p.id} style={{ backgroundColor: i%2===0?C.white:C.tableStripe }}>
                      <td style={{ ...tR, fontFamily:"monospace", fontSize:11, color:C.textMuted }}>{p.jobNo}</td>
                      <td style={{ ...tR, fontWeight:600, color:C.textPrimary }}>{p.jobName}</td>
                      <td style={{ ...tR, color:C.textMuted }}>{p.client||"—"}</td>
                      <td style={{ ...tR, color:C.textMuted }}>{p.location||"—"}</td>
                      <td style={{ ...tR, color:C.textMuted }}>{p.bidType||"—"}</td>
                      <td style={{ ...tR, fontFamily:"monospace" }}>{p.dueDate}</td>
                      <td style={{ ...tR, textAlign:"center" }}>
                        <span style={{ padding:"2px 8px", borderRadius:10, fontSize:11, fontWeight:700, backgroundColor:uc.bg, color:uc.text }}>
                          {p.daysOut === 0 ? "Today" : `${p.daysOut}d`}
                        </span>
                      </td>
                      <td style={tR}>
                        <span style={{ padding:"2px 8px", borderRadius:3, fontSize:11, fontWeight:600, ...(STATUS_STYLES[p.status]||{}) }}>
                          {p.status}
                        </span>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}

        {/* ── Middle row: Industry + Bid Type + Status ── */}
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:16 }}>

          {/* Industry */}
          <div style={{ backgroundColor:C.white, border:`1px solid ${C.border}`, borderRadius:4 }}>
            <div style={{ padding:"12px 16px 8px", borderBottom:`1px solid ${C.border}`, fontSize:13, fontWeight:700, color:C.textPrimary }}>🏗️ Top Industries</div>
            <table style={{ width:"100%", borderCollapse:"collapse" }}>
              <thead><tr>
                {["Industry","Won","Lost","Rev Won","Hit %"].map(h=><th key={h} style={{ ...tH, fontSize:10 }}>{h}</th>)}
              </tr></thead>
              <tbody>
                {byIndustry.length===0
                  ? <tr><td colSpan={5} style={{ ...tR, color:C.textMuted, fontStyle:"italic", textAlign:"center" }}>No data</td></tr>
                  : byIndustry.map((r,i)=>(
                  <tr key={r.ind} style={{ backgroundColor:i%2===0?C.white:C.tableStripe }}>
                    <td style={tR}>{r.ind}</td>
                    <td style={{ ...tR, textAlign:"center", color:C.success, fontWeight:600 }}>{r.won}</td>
                    <td style={{ ...tR, textAlign:"center", color:C.danger }}>{r.lost}</td>
                    <td style={{ ...tRr, fontSize:11 }}>{r.revWon>0?fmtM(r.revWon):"—"}</td>
                    <td style={{ ...tRr, fontWeight:600, color: r.hitRate>=0.5?C.success:C.danger }}>{fmtPct(r.hitRate)}</td>
                  </tr>
                ))}
                {byIndustry.length>0 && (
                  <tr style={{ backgroundColor:"#F0F4FF", fontWeight:700 }}>
                    <td style={tR}>TOTAL</td>
                    <td style={{ ...tR, textAlign:"center", color:C.success }}>{metrics.jobsWon}</td>
                    <td style={{ ...tR, textAlign:"center", color:C.danger }}>{byStatus.find(s=>s.status==="Lost")?.count||0}</td>
                    <td style={{ ...tRr, fontSize:11 }}>{fmtM(metrics.revenueWon)}</td>
                    <td style={{ ...tRr, fontWeight:700 }}>{fmtPct(metrics.hitRate)}</td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>

          {/* Bid Type */}
          <div style={{ backgroundColor:C.white, border:`1px solid ${C.border}`, borderRadius:4 }}>
            <div style={{ padding:"12px 16px 8px", borderBottom:`1px solid ${C.border}`, fontSize:13, fontWeight:700, color:C.textPrimary }}>⚖️ Hard Bid vs. Negotiated</div>
            <table style={{ width:"100%", borderCollapse:"collapse" }}>
              <thead><tr>
                {["Bid Type","Won","Lost","Rev Won","% of Won"].map(h=><th key={h} style={{ ...tH, fontSize:10 }}>{h}</th>)}
              </tr></thead>
              <tbody>
                {byBidType.length===0
                  ? <tr><td colSpan={5} style={{ ...tR, color:C.textMuted, fontStyle:"italic", textAlign:"center" }}>No data</td></tr>
                  : byBidType.map((r,i)=>(
                  <tr key={r.type} style={{ backgroundColor:i%2===0?C.white:C.tableStripe }}>
                    <td style={tR}>{r.type}</td>
                    <td style={{ ...tR, textAlign:"center", color:C.success, fontWeight:600 }}>{r.won}</td>
                    <td style={{ ...tR, textAlign:"center", color:C.danger }}>{r.lost}</td>
                    <td style={{ ...tRr, fontSize:11 }}>{r.revWon>0?fmtM(r.revWon):"—"}</td>
                    <td style={{ ...tRr }}>{metrics.revenueWon>0?fmtPct(r.revWon/metrics.revenueWon):"—"}</td>
                  </tr>
                ))}
                {byBidType.length>0 && (
                  <tr style={{ backgroundColor:"#F0F4FF", fontWeight:700 }}>
                    <td style={tR}>TOTAL</td>
                    <td style={{ ...tR, textAlign:"center", color:C.success }}>{byBidType.reduce((a,r)=>a+r.won,0)}</td>
                    <td style={{ ...tR, textAlign:"center", color:C.danger }}>{byBidType.reduce((a,r)=>a+r.lost,0)}</td>
                    <td style={{ ...tRr, fontSize:11 }}>{fmtM(metrics.revenueWon)}</td>
                    <td style={{ ...tRr }}>100%</td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>

          {/* Status Breakdown */}
          <div style={{ backgroundColor:C.white, border:`1px solid ${C.border}`, borderRadius:4 }}>
            <div style={{ padding:"12px 16px 8px", borderBottom:`1px solid ${C.border}`, fontSize:13, fontWeight:700, color:C.textPrimary }}>📊 Bid Status Breakdown</div>
            <table style={{ width:"100%", borderCollapse:"collapse" }}>
              <thead><tr>
                {["Status","Count","% of Total","Revenue","Proj. Fee"].map(h=><th key={h} style={{ ...tH, fontSize:10 }}>{h}</th>)}
              </tr></thead>
              <tbody>
                {byStatus.map((r,i)=>(
                  r.count===0 ? null :
                  <tr key={r.status} style={{ backgroundColor:i%2===0?C.white:C.tableStripe }}>
                    <td style={tR}>
                      <span style={{ padding:"2px 8px", borderRadius:3, fontSize:11, fontWeight:600, ...(STATUS_STYLES[r.status]||{}) }}>{r.status}</span>
                    </td>
                    <td style={{ ...tR, textAlign:"center", fontWeight:600 }}>{r.count}</td>
                    <td style={{ ...tRr }}>{metrics.totalBids>0?fmtPct(r.count/metrics.totalBids):"—"}</td>
                    <td style={{ ...tRr, fontSize:11 }}>{r.rev>0?fmtM(r.rev):"—"}</td>
                    <td style={{ ...tRr, fontSize:11 }}>{r.fee>0?fmtM(r.fee):"—"}</td>
                  </tr>
                ))}
                <tr style={{ backgroundColor:"#F0F4FF", fontWeight:700 }}>
                  <td style={tR}>TOTAL</td>
                  <td style={{ ...tR, textAlign:"center" }}>{metrics.totalBids}</td>
                  <td style={{ ...tRr }}>100%</td>
                  <td style={{ ...tRr, fontSize:11 }}>{fmtM(byStatus.reduce((a,r)=>a+r.rev,0))}</td>
                  <td style={{ ...tRr, fontSize:11 }}>{fmtM(byStatus.reduce((a,r)=>a+r.fee,0))}</td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>

        {/* ── Bottom row: Top Clients + YoY ── */}
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:16 }}>

          {/* Top Clients */}
          <div style={{ backgroundColor:C.white, border:`1px solid ${C.border}`, borderRadius:4 }}>
            <div style={{ padding:"12px 16px 8px", borderBottom:`1px solid ${C.border}`, fontSize:13, fontWeight:700, color:C.textPrimary }}>🤝 Top Clients by Revenue Won</div>
            <table style={{ width:"100%", borderCollapse:"collapse" }}>
              <thead><tr>
                {["Client","Jobs Won","Revenue Won","Proj. Fee"].map(h=><th key={h} style={{ ...tH, fontSize:10 }}>{h}</th>)}
              </tr></thead>
              <tbody>
                {topClients.length===0
                  ? <tr><td colSpan={4} style={{ ...tR, color:C.textMuted, fontStyle:"italic", textAlign:"center" }}>No won projects yet</td></tr>
                  : topClients.map((r,i)=>(
                  <tr key={r.client} style={{ backgroundColor:i%2===0?C.white:C.tableStripe }}>
                    <td style={{ ...tR, fontWeight:500 }}>{r.client}</td>
                    <td style={{ ...tR, textAlign:"center", fontWeight:600, color:C.success }}>{r.won}</td>
                    <td style={{ ...tRr }}>{fmtRev(r.rev)}</td>
                    <td style={{ ...tRr, color:C.textMuted }}>{fmtRev(r.fee)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {/* Year-over-Year */}
          <div style={{ backgroundColor:C.white, border:`1px solid ${C.border}`, borderRadius:4 }}>
            <div style={{ padding:"12px 16px 8px", borderBottom:`1px solid ${C.border}`, fontSize:13, fontWeight:700, color:C.textPrimary }}>📅 Year-over-Year Summary</div>
            <table style={{ width:"100%", borderCollapse:"collapse" }}>
              <thead><tr>
                {["Year","Total Bids","Won","Lost","Revenue Won","Hit Rate"].map(h=><th key={h} style={{ ...tH, fontSize:10 }}>{h}</th>)}
              </tr></thead>
              <tbody>
                {yoy.length===0
                  ? <tr><td colSpan={6} style={{ ...tR, color:C.textMuted, fontStyle:"italic", textAlign:"center" }}>No data yet</td></tr>
                  : yoy.map((r,i)=>(
                  <tr key={r.yr} style={{ backgroundColor:i%2===0?C.white:C.tableStripe }}>
                    <td style={{ ...tR, fontWeight:700, fontFamily:"monospace" }}>{r.yr}</td>
                    <td style={{ ...tR, textAlign:"center" }}>{r.total}</td>
                    <td style={{ ...tR, textAlign:"center", fontWeight:600, color:C.success }}>{r.won}</td>
                    <td style={{ ...tR, textAlign:"center", color:C.danger }}>{r.lost}</td>
                    <td style={{ ...tRr }}>{r.revWon>0?fmtRev(r.revWon):"—"}</td>
                    <td style={{ ...tRr, fontWeight:700, color:r.hitRate>=0.5?C.success:r.hitRate!=null?C.danger:C.textMuted }}>{fmtPct(r.hitRate)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        {/* Empty state */}
        {allProjects.length === 0 && (
          <div style={{ textAlign:"center", padding:"60px 0", color:C.textMuted }}>
            <div style={{ fontSize:40, marginBottom:12 }}>📊</div>
            <div style={{ fontSize:15, fontWeight:600, color:C.textPrimary, marginBottom:6 }}>No bid data yet</div>
            <div style={{ fontSize:13 }}>Create estimates in the Estimates module to start tracking your bid pipeline.</div>
          </div>
        )}
      </div>
    </div>
  );
}

// ─── App Shell ────────────────────────────────────────────────────────────────
export default function App() {
  const [activeNav, setActiveNav] = useState("cost-database");

  const pageTitle = NAV_ITEMS.find(n => n.key === activeNav)?.label ?? "";

  return (
    <div style={s.shell}>
      {/* Nav */}
      <nav style={s.nav}>
        <div style={s.navHeader}>
          <div style={s.navOrgName}>{MOCK_ORG.name}</div>
          <div style={s.navOrgPlan}>{MOCK_ORG.plan}</div>
        </div>
        <div style={s.navSection}>
          <div style={s.navSectionLabel}>Workspace</div>
          {NAV_ITEMS.map(item => (
            <div key={item.key} style={s.navItem(activeNav === item.key)} onClick={() => setActiveNav(item.key)}>
              <span style={s.navDot(activeNav === item.key)} />
              {item.label}
            </div>
          ))}
        </div>
        <div style={{ ...s.navSection, marginTop: 8 }}>
          <div style={s.navSectionLabel}>Settings</div>
          {["Org Settings", "User Management", "Lookup Values"].map(label => {
            const key = label.toLowerCase().replace(/ /g, "-");
            return (
              <div key={key} style={s.navItem(activeNav === key)} onClick={() => setActiveNav(key)}>
                <span style={s.navDot(activeNav === key)} />
                {label}
              </div>
            );
          })}
        </div>
        <div style={s.navFooter}>
          <div style={s.avatar}>{MOCK_USER.initials}</div>
          <div>
            <div style={s.navUserName}>{MOCK_USER.name}</div>
            <div style={s.navUserRole}>{MOCK_USER.role}</div>
          </div>
        </div>
      </nav>

      {/* Main */}
      <div style={s.main}>
        <div style={s.topBar}>
          <span style={s.topBarTitle}>{pageTitle}</span>
          <span style={{ fontSize: 12, color: C.textMuted }}>{MOCK_ORG.name}</span>
        </div>

        {activeNav === "cost-database" ? (
          <CostDatabase />
        ) : activeNav === "estimates" ? (
          <Estimates />
        ) : activeNav === "dashboard" ? (
          <Dashboard />
        ) : activeNav === "bid-comparisons" ? (
          <BidComparison />
        ) : (
          <div style={{ flex: 1, display: "flex", alignItems: "center", justifyContent: "center", color: C.textMuted, fontSize: 13 }}>
            <div style={{ textAlign: "center" }}>
              <div style={{ fontSize: 28, marginBottom: 8 }}>⬡</div>
              <div style={{ fontWeight: 500, marginBottom: 4 }}>{pageTitle}</div>
              <div>This module is on the roadmap.</div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
