import { Router } from 'express';
import {
  getCandidateByExamNumber,
  listCandidates,
  createCandidate,
  updateCandidate,
  deleteCandidate
} from '../controllers/candidate.controller.js';
import { protectAdmin } from '../middleware/auth.js';

const router = Router();

router.get('/lookup/:examNumber', getCandidateByExamNumber);
router.get('/', protectAdmin, listCandidates);
router.post('/', protectAdmin, createCandidate);
router.put('/:id', protectAdmin, updateCandidate);
router.delete('/:id', protectAdmin, deleteCandidate);

export default router;
