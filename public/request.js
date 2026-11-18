const form = document.querySelector('#request-form');
const message = document.querySelector('#form-message');
form.addEventListener('submit', async (event) => {
  event.preventDefault();
  const button = form.querySelector('button');
  button.disabled = true; button.textContent = 'Sending...';
  const data = Object.fromEntries(new FormData(form).entries());
  try {
    const response = await fetch('/api/requests', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(data) });
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || 'Unable to submit request.');
    form.hidden = true; message.hidden = false; message.className = 'message success';
    message.innerHTML = '<strong>Thank you — your request has been received.</strong><p>Our team will review the information and contact you. Your request number is #' + Number(result.id) + '.</p>';
  } catch (error) {
    message.hidden = false; message.className = 'message error-box'; message.textContent = error.message;
    button.disabled = false; button.textContent = 'Send Estimate Request →';
  }
});