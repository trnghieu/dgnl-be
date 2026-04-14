import Candidate from '../models/Candidate.js';
import { SUBJECTS, SUBJECT_KEYS } from '../constants/subjects.js';

const sanitizeScores = (scores = {}, { preserveMissing = false } = {}) =>
  SUBJECT_KEYS.reduce((accumulator, key) => {
    const value = scores[key];

    if (preserveMissing && value === undefined) {
      return accumulator;
    }

    if (value === '' || value === undefined || value === null) {
      accumulator[key] = null;
      return accumulator;
    }

    const parsed = Number(value);
    accumulator[key] = Number.isFinite(parsed) ? parsed : null;
    return accumulator;
  }, {});

export const getCandidateByExamNumber = async (req, res, next) => {
  try {
    const candidate = await Candidate.findOne({ examNumber: req.params.examNumber.trim() });

    if (!candidate) {
      return res.status(404).json({ message: 'Không tìm thấy thí sinh.' });
    }

    return res.json({
      candidate,
      subjects: SUBJECTS
    });
  } catch (error) {
    next(error);
  }
};

export const listCandidates = async (req, res, next) => {
  try {
    const { page = 1, limit = 300, search = '' } = req.query;
    const query = search
      ? {
          $or: [
            { examNumber: { $regex: search, $options: 'i' } },
            { fullName: { $regex: search, $options: 'i' } },
            { className: { $regex: search, $options: 'i' } }
          ]
        }
      : {};

    const [items, total] = await Promise.all([
      Candidate.find(query)
        .sort({ createdAt: -1 })
        .skip((Number(page) - 1) * Number(limit))
        .limit(Number(limit)),
      Candidate.countDocuments(query)
    ]);

    res.json({
      items,
      pagination: {
        page: Number(page),
        limit: Number(limit),
        total,
        totalPages: Math.ceil(total / Number(limit)) || 1
      },
      subjects: SUBJECTS
    });
  } catch (error) {
    next(error);
  }
};

export const createCandidate = async (req, res, next) => {
  try {
    const { examNumber, fullName, birthDate, className, scores } = req.body;

    if (!examNumber || !fullName) {
      res.status(400);
      throw new Error('examNumber và fullName là bắt buộc.');
    }

    const candidate = await Candidate.create({
      examNumber: String(examNumber).trim(),
      fullName: String(fullName).trim(),
      birthDate: birthDate || '',
      className: className || '',
      scores: sanitizeScores(scores)
    });

    res.status(201).json({ message: 'Tạo thí sinh thành công.', candidate });
  } catch (error) {
    next(error);
  }
};

export const updateCandidate = async (req, res, next) => {
  try {
    const { examNumber, fullName, birthDate, className, scores } = req.body;

    const candidate = await Candidate.findById(req.params.id);

    if (!candidate) {
      return res.status(404).json({ message: 'Không tìm thấy thí sinh.' });
    }

    candidate.examNumber = examNumber ? String(examNumber).trim() : candidate.examNumber;
    candidate.fullName = fullName ? String(fullName).trim() : candidate.fullName;
    candidate.birthDate = birthDate ?? candidate.birthDate;
    candidate.className = className ?? candidate.className;
    candidate.scores = { ...candidate.scores.toObject(), ...sanitizeScores(scores, { preserveMissing: true }) };

    await candidate.save();

    res.json({ message: 'Cập nhật thành công.', candidate });
  } catch (error) {
    next(error);
  }
};

export const deleteCandidate = async (req, res, next) => {
  try {
    const candidate = await Candidate.findById(req.params.id);

    if (!candidate) {
      return res.status(404).json({ message: 'Không tìm thấy thí sinh.' });
    }

    await candidate.deleteOne();
    res.json({ message: 'Xóa thí sinh thành công.' });
  } catch (error) {
    next(error);
  }
};
