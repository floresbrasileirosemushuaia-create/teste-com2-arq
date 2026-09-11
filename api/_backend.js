'use strict';
const VERSION='2.4.4';
const FALLBACK='https://script.google.com/macros/s/AKfycbx9H9BJTrwNXdORIrVwilUmQiCMq-enVm-HfXZjNjXKsudiklef9ZnMdc36ZtIsD7bB/exec';
function backendUrl(){const raw=String(process.env.B2B_BACKEND_URL||FALLBACK||'').trim();if(!/^https:\/\/script\.google\.com\/macros\/s\/[A-Za-z0-9_-]+\/exec$/.test(raw))return '';return raw}
module.exports={VERSION,backendUrl};
