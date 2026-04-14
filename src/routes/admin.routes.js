import { Router } from 'express';
import multer from 'multer';
import { protectAdmin } from '../middleware/auth.js';
import {
  importCandidatesFromExcel,
  exportCandidatesToExcel,
  dashboardSummary
} from '../controllers/admin.controller.js';

const router = Router();
const upload = multer({ storage: multer.memoryStorage() });

router.use(protectAdmin);
router.get('/summary', dashboardSummary);
router.post('/import', upload.single('file'), importCandidatesFromExcel);
router.get('/export', exportCandidatesToExcel);

export default router;
