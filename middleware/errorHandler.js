const errorHandler = (err, req, res, next) => {
  console.error('Error:', err);

  const statusCode = err.statusCode || 500;
  const message = err.message || 'Internal Server Error';

  // Convert error response to match frontend expectations
  const errorResponse = {
    success: false,
    error: message,
    details: process.env.NODE_ENV === 'development' ? err.stack : undefined
  };

  // Send as JSON instead of blob
  res.status(statusCode).json(errorResponse);
};

module.exports = errorHandler;