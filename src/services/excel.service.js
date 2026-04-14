import XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import Candidate from '../models/Candidate.js';

const SUBJECT_LABELS = {
  math: ['Toán', 'math'],
  physics: ['Lý', 'physics'],
  chemistry: ['Hóa', 'chemistry']
};

const FIELD_LABELS = {
  examNumber: ['Số báo danh', 'examNumber'],
  fullName: ['Họ và tên', 'fullName'],
  birthDate: ['Ngày sinh', 'birthDate'],
  className: ['Lớp', 'className']
};

const SUBJECT_KEYS = ['math', 'physics', 'chemistry'];

const getCellValue = (row, possibleKeys) => {
  for (const key of possibleKeys) {
    if (row[key] !== undefined && row[key] !== null && row[key] !== '') {
      return row[key];
    }
  }
  return '';
};

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
    const examNumber = String(getCellValue(row, FIELD_LABELS.examNumber)).trim();
    const fullName = String(getCellValue(row, FIELD_LABELS.fullName)).trim();
    const birthDate = String(getCellValue(row, FIELD_LABELS.birthDate)).trim();
    const className = String(getCellValue(row, FIELD_LABELS.className)).trim();

    if (!examNumber || !fullName) {
      throw new Error(`Dòng ${index + 2} thiếu Số báo danh hoặc Họ và tên.`);
    }

    const scores = SUBJECT_KEYS.reduce((accumulator, subjectKey) => {
      accumulator[subjectKey] = normalizeScore(
        getCellValue(row, SUBJECT_LABELS[subjectKey])
      );
      return accumulator;
    }, {});

    return {
      examNumber,
      fullName,
      birthDate,
      className,
      scores
    };
  });
};

export const buildCandidatesWorkbookBuffer = async () => {
  const candidates = await Candidate.find().sort({ examNumber: 1 }).lean();
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Danh sách điểm');

  worksheet.columns = [
    { header: 'Số báo danh', key: 'examNumber', width: 18 },
    { header: 'Họ và tên', key: 'fullName', width: 28 },
    { header: 'Ngày sinh', key: 'birthDate', width: 16 },
    { header: 'Lớp', key: 'className', width: 14 },
    { header: 'Toán', key: 'math', width: 12 },
    { header: 'Lý', key: 'physics', width: 12 },
    { header: 'Hóa', key: 'chemistry', width: 12 },
    { header: 'Tổng điểm', key: 'totalScore', width: 14 }
  ];

  candidates.forEach((candidate) => {
    const math = typeof candidate.scores?.math === 'number' ? candidate.scores.math : '';
    const physics =
      typeof candidate.scores?.physics === 'number' ? candidate.scores.physics : '';
    const chemistry =
      typeof candidate.scores?.chemistry === 'number' ? candidate.scores.chemistry : '';

    const totalScore =
      (typeof candidate.scores?.math === 'number' ? candidate.scores.math : 0) +
      (typeof candidate.scores?.physics === 'number' ? candidate.scores.physics : 0) +
      (typeof candidate.scores?.chemistry === 'number'
        ? candidate.scores.chemistry
        : 0);

    worksheet.addRow({
      examNumber: candidate.examNumber,
      fullName: candidate.fullName,
      birthDate: candidate.birthDate || '',
      className: candidate.className || '',
      math,
      physics,
      chemistry,
      totalScore
    });
  });

  worksheet.getRow(1).font = { bold: true };
  worksheet.views = [{ state: 'frozen', ySplit: 1 }];

  return workbook.xlsx.writeBuffer();
};