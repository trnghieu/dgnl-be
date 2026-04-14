import Candidate from '../models/Candidate.js';
import { parseCandidatesFromExcelBuffer, buildCandidatesWorkbookBuffer } from '../services/excel.service.js';

export const importCandidatesFromExcel = async (req, res, next) => {
  try {
    if (!req.file) {
      res.status(400);
      throw new Error('Vui lòng tải lên file Excel.');
    }

    const rows = parseCandidatesFromExcelBuffer(req.file.buffer);
    let inserted = 0;
    let updated = 0;

    for (const row of rows) {
      const existingCandidate = await Candidate.findOne({ examNumber: row.examNumber });

      if (existingCandidate) {
        existingCandidate.fullName = row.fullName;
        existingCandidate.birthDate = row.birthDate;
        existingCandidate.className = row.className;
        existingCandidate.scores = row.scores;
        await existingCandidate.save();
        updated += 1;
      } else {
        await Candidate.create(row);
        inserted += 1;
      }
    }

    res.json({
      message: 'Import Excel thành công.',
      summary: {
        totalRows: rows.length,
        inserted,
        updated
      }
    });
  } catch (error) {
    next(error);
  }
};

export const exportCandidatesToExcel = async (req, res, next) => {
  try {
    const buffer = await buildCandidatesWorkbookBuffer();

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="candidates.xlsx"');
    res.send(Buffer.from(buffer));
  } catch (error) {
    next(error);
  }
};

export const dashboardSummary = async (req, res, next) => {
  try {
    const totalCandidates = await Candidate.countDocuments();
    res.json({ totalCandidates });
  } catch (error) {
    next(error);
  }
};
