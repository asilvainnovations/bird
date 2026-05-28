require('dotenv').config();

module.exports = {
  // --- Color Palette ---
  COLORS: {
    primary:    '1B3A6B',   // Deep Navy
    secondary:  '2E7D32',   // Green
    accent:     'C0392B',   // Red
    gold:       'D4A017',   // Bangsamoro Gold
    lightBlue:  'D6E8F5',
    lightGreen: 'D5EDD5',
    lightGold:  'FFF3CD',
    lightGray:  'F5F5F5',
    midGray:    'BBBBBB',
    white:      'FFFFFF',
  },

  // --- Document Dimensions (in DXA units - 1/1440 inch) ---
  FULL_W: 9360,           // Content width (8.5" - 2" margins)
  PAGE_HEIGHT: 15840,     // A4 height in DXA
  MARGIN: 1440,           // 1" margin in all sides

  // --- File Paths ---
  OUTPUT_DIR: './outputs',
  DOCX_PATH: `${process.env.OUTPUT_DIR}/BIRD_2026_2035_Bangsamoro_Investment_Roadmap.docx`,
  PDF_PATH: `${process.env.OUTPUT_DIR}/pdfs/BIRD_2026_2035_Bangsamoro_Investment_Roadmap.pdf`,
  
  // --- Google Drive Settings ---
  GOOGLE_DRIVE: {
    SERVICE_ACCOUNT_KEY: process.env.GOOGLE_SERVICE_ACCOUNT_KEY || './credentials/google-key.json',
    FOLDER_ID: process.env.GOOGLE_FOLDER_ID || '',
    SHARED_WITH: ['mtit-barmm@example.com'] // Add distribution list
  },

  // --- Version Control ---
  VERSION: '1.0.0',
  DATE: new Date().toISOString().split('T')[0],
  TITLE: 'BANGSAMORO INVESTMENT ROADMAP DEVELOPMENT (BIRD) 2026–2035',
  SUBTITLE: 'Revealing the Bangsamoro\'s Hidden Treasure: An Integrated, Ecosystem-Based Investment Strategy',

  // --- Contact Information ---
  CONTACT: {
    ministry: 'Ministry of Trade, Investments and Tourism (MTIT)',
    bureau: 'Bureau of Investments (BOI-MTIT)',
    city: 'Cotabato City, Bangsamoro',
    email: 'investments@barmm.gov.ph',
    website: 'https://mtit.bangsamoro.gov.ph'
  },

  // --- Security Settings ---
  SECURITY: {
    PASSWORD: 'BARMm2026Invest!',
    READONLY_PASSWORD: 'SecureView2026!'
  }
};
