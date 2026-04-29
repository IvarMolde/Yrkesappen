'use strict';

function errorHandler(err, req, res, next) {
  console.error('Uventet serverfeil:', err);

  if (res.headersSent) {
    return next(err);
  }

  return res.status(500).json({
    feil: 'En uventet feil oppstod. Prøv igjen om litt.',
  });
}

module.exports = { errorHandler };
