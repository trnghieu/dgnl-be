import Admin from '../models/Admin.js';
import { generateToken } from '../utils/generateToken.js';

export const loginAdmin = async (req, res, next) => {
  try {
    const { username, password } = req.body;

    if (!username || !password) {
      res.status(400);
      throw new Error('Vui lòng nhập username và password.');
    }

    const admin = await Admin.findOne({ username });

    if (!admin || !(await admin.comparePassword(password))) {
      res.status(401);
      throw new Error('Sai tài khoản hoặc mật khẩu.');
    }

    const token = generateToken(admin);

    res.json({
      message: 'Đăng nhập thành công.',
      token,
      admin: {
        id: admin._id,
        username: admin.username,
        fullName: admin.fullName
      }
    });
  } catch (error) {
    next(error);
  }
};
