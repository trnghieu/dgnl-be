export const notFoundHandler = (req, res) => {
  res.status(404).json({ message: 'API không tồn tại.' });
};

export const errorHandler = (error, req, res, next) => {
  const statusCode = res.statusCode && res.statusCode !== 200 ? res.statusCode : 500;

  if (error?.code === 11000) {
    return res.status(400).json({ message: 'Số báo danh đã tồn tại.' });
  }

  return res.status(statusCode).json({
    message: error.message || 'Đã xảy ra lỗi trên server.',
    stack: process.env.NODE_ENV === 'production' ? undefined : error.stack
  });
};
