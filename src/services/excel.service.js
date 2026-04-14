import XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import Candidate from '../models/Candidate.js';
import { SUBJECTS, SUBJECT_KEYS } from '../constants/subjects.js';

const normalizeScore = (value) => {
  if (value === '' || value === undefined || value === null) {
    return null;
  }

  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : null;
};

export const parseCandidatesFromExcelBuffer = (buffer) => {
  const workbook = XLSX.read(buffer, { type: 'buffer' });
  const firstSheet = workbook.SheetNames[0];

  if (!firstSheet) {
    throw new Error('File Excel không có sheet dữ liệu.');
  }

  const rawRows = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], {
    defval: ''
  });

  if (!rawRows.length) {
    throw new Error('File Excel không có dữ liệu.');
  }

  return rawRows.map((row, index) => {
    if (!row.examNumber || !row.fullName) {
      throw new Error(`Dòng ${index + 2} thiếu examNumber hoặc fullName.`);
    }

    const scores = SUBJECT_KEYS.reduce((accumulator, subjectKey) => {
      accumulator[subjectKey] = normalizeScore(row[subjectKey]);
      return accumulator;
    }, {});

    return {
      examNumber: String(row.examNumber).trim(),
      fullName: String(row.fullName).trim(),
      birthDate: row.birthDate ? String(row.birthDate).trim() : '',
      className: row.className ? String(row.className).trim() : '',
      scores
    };
  });
};

export const buildCandidatesWorkbookBuffer = async () => {
  const candidates = await Candidate.find().sort({ examNumber: 1 }).lean();
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Candidates');

  const columns = [
    { header: 'examNumber', key: 'examNumber', width: 18 },
    { header: 'fullName', key: 'fullName', width: 28 },
    { header: 'birthDate', key: 'birthDate', width: 14 },
    { header: 'className', key: 'className', width: 16 },
    ...SUBJECTS.map((subject) => ({
      header: subject.key,
      key: subject.key,
      width: 14
    })),
    { header: 'totalScore', key: 'totalScore', width: 14 }
  ];

  worksheet.columns = columns;

  candidates.forEach((candidate) => {
    const row = {
      examNumber: candidate.examNumber,
      fullName: candidate.fullName,
      birthDate: candidate.birthDate,
      className: candidate.className,
      totalScore: SUBJECT_KEYS.reduce((total, key) => {
        const value = candidate.scores?.[key];
        return typeof value === 'number' ? total + value : total;
      }, 0)
    };

    SUBJECT_KEYS.forEach((key) => {
      row[key] = candidate.scores?.[key] ?? '';
    });

    worksheet.addRow(row);
  });

  worksheet.getRow(1).font = { bold: true };
  worksheet.views = [{ state: 'frozen', ySplit: 1 }];

  return workbook.xlsx.writeBuffer();
};
