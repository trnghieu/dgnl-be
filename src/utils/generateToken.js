import jwt from 'jsonwebtoken';
import { env } from '../config/env.js';

export const generateToken = (admin) =>
  jwt.sign(
    {
      id: admin._id,
      username: admin.username
    },
    env.jwtSecret,
    { expiresIn: env.jwtExpiresIn }
  );
