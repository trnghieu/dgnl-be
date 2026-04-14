import mongoose from 'mongoose';
import { env } from '../config/env.js';
import Admin from '../models/Admin.js';

const seedAdmin = async () => {
  try {
    await mongoose.connect(env.mongodbUri);

    const existing = await Admin.findOne({ username: env.adminSeedUsername });

    if (existing) {
      existing.password = env.adminSeedPassword;
      await existing.save();
      console.log('Admin đã được cập nhật lại mật khẩu.');
    } else {
      await Admin.create({
        username: env.adminSeedUsername,
        password: env.adminSeedPassword,
        fullName: 'Administrator'
      });
      console.log('Tạo admin thành công.');
    }
  } catch (error) {
    console.error('Seed admin thất bại:', error);
  } finally {
    await mongoose.disconnect();
  }
};

seedAdmin();
