let requests = [];
let selected = null;
const $ = (s) => document.querySelector(s);
const money = (n) => new Intl.NumberFormat('en-US',{style:'currency',currency:'USD'}).format(Number(n||0));
async function api(url, options={}) {
  const response = await fetch(url,{...options,headers:{'Content-Type':'application/json',...(options.headers||{})}});
  const data = await response.json();
  if (!response.ok) throw new Error(data.error || 'Something went wrong.');
  return data;
}
function showApp(signedIn) { $('#login-panel').hidden=signedIn; $('#app-panel').hidden=!signedIn; }
async function checkEmailConnection(){try{const r=await api('/api/email/status');$('#outlook-status').textContent=r.connected?'Outlook is connected. Approved estimates will be sent from the connected account.':'Outlook is not connected yet.';$('#outlook-connect').textContent=r.connected?'Reconnect Outlook':'Connect Outlook';}catch(e){$('#outlook-status').textContent='Outlook connection setup is not configured yet.';}}\nasync function checkLogin() { try { const r=await api('/api/me'); showApp(r.signedIn); if(r.signedIn){await loadRequests();await checkEmailConnection();} } catch(e) { showApp(false); } }
$('#login-form').addEventListener('submit',async e=>{e.preventDefault();$('#login-error').textContent='';try{await api('/api/login',{method:'POST',body:JSON.stringify({password:new FormData(e.target).get('password')})});e.target.reset();showApp(true);await loadRequests();}catch(err){$('#login-error').textContent=err.message;}});
$('#logout').addEventListener('click',async()=>{await api('/api/logout',{method:'POST'});showApp(false);});
$('#refresh').addEventListener('click',loadRequests);
async function loadRequests(){
  try { requests=await api('/api/requests'); renderStats(); renderList(); if(selected){const current=requests.find(r=>String(r.id)===String(selected.id)); if(current){selected=current;renderDetail();}else clearDetail();} }
  catch(e){$('#request-list').textContent=e.message;}
}
function renderStats(){
 const counts={New:0,'In Review':0,'Estimate Draft':0,Approved:0,Sent:0};
 requests.forEach(r=>{if(counts[r.status]!==undefined)counts[r.status]++;});
 $('#stats').innerHTML=[['New Requests',counts.New],['In Review',counts['In Review']],['Draft Estimates',counts['Estimate Draft']],['Approved',counts.Approved],['Emailed',counts.Sent]].map(x=>'<div class="stat-card"><span>'+x[0]+'</span><strong>'+x[1]+'</strong></div>').join('');
}
function renderList(){
 if(!requests.length){$('#request-list').innerHTML='<div class="empty-list">No estimate requests yet. New customer submissions will appear here.</div>';return;}
 $('#request-list').innerHTML=requests.map(r=>'<button class="request-row '+(selected&&String(selected.id)===String(r.id)?'active':'')+'" data-id="'+r.id+'"><span class="request-top"><strong>#'+r.id+' · '+esc(r.customer_name)+'</strong><span class="status '+slug(r.status)+'">'+esc(r.status)+'</span></span><span class="request-vehicle">'+esc([r.year,r.make,r.model].filter(Boolean).join(' '))+'</span><span class="request-bottom">'+esc(r.vehicle_type)+' · '+new Date(r.created_at).toLocaleDateString()+'</span></button>').join('');
 document.querySelectorAll('.request-row').forEach(b=>b.addEventListener('click',()=>{selected=requests.find(r=>String(r.id)===b.dataset.id);renderList();renderDetail();}));
}
function clearDetail(){selected=null;$('#empty-detail').hidden=false;$('#request-detail').hidden=true;}
function renderDetail(){
 if(!selected)return;$('#empty-detail').hidden=true;$('#request-detail').hidden=false;
 const r=selected, estimate=r.estimate||{items:[],customerMessage:''};
 const items=estimate.items||[];
 const emailStatus=r.estimate_email_sent_at?'<p class="fine-print">Estimate emailed to '+esc(r.email||'customer')+' on '+new Date(r.estimate_email_sent_at).toLocaleString()+'.</p>':(r.estimate_email_error?'<p class="error">Email not sent: '+esc(r.estimate_email_error)+'</p>':'');
 const customerStatus=r.customer_decision?'<p class="fine-print">Customer response: <strong>'+esc(r.customer_decision)+'</strong> · '+new Date(r.customer_decision_at).toLocaleString()+'</p>':'';
 $('#request-detail').innerHTML='<div class="detail-head"><div><p class="eyebrow">REQUEST #'+r.id+'</p><h2>'+esc(r.customer_name)+'</h2><p class="muted">'+esc(r.phone)+(r.email?' · '+esc(r.email):'')+'</p></div><select id="status-select" aria-label="Request status">'+['New','In Review','Estimate Draft','Approved','Sent','Declined','Completed'].map(s=>'<option '+(r.status===s?'selected':'')+'>'+s+'</option>').join('')+'</select></div>'+
 '<div class="vehicle-box"><p class="eyebrow">VEHICLE</p><h3>'+esc([r.year,r.make,r.model].filter(Boolean).join(' '))+'</h3><p>'+esc(r.vehicle_type)+' · '+esc(r.engine||'Engine not specified')+'</p><p>VIN: '+esc(r.vin||'Not provided')+' · Mileage: '+esc(r.mileage||'Not provided')+'</p></div>'+
 '<div class="concern-box"><p class="eyebrow">CUSTOMER CONCERN</p><p>'+esc(r.concern)+'</p></div>'+
 '<div class="estimate-editor"><div class="panel-heading"><div><p class="eyebrow">OWNER WORKSPACE</p><h3>Build Estimate</h3></div><span class="rate-note">$135/hr · 40% parts markup</span></div>'+
 '<div class="estimate-options"><label>Parts preference<select id="parts-type"><option>OEM</option><option>Aftermarket</option><option>Customer Supplied</option></select></label><label>Customer message<input id="customer-message" value="'+esc(estimate.customerMessage||'')+'" placeholder="Optional note for customer"></label></div>'+
 '<div id="line-items">'+items.map(itemRow).join('')+'</div><button id="add-line" class="button secondary small">+ Add line item</button>'+
 '<div class="totals" id="totals"></div><label>Internal notes<textarea id="owner-notes" rows="2">'+esc(r.owner_notes||'')+'</textarea></label>'+
 '<div class="action-row"><button id="save-draft" class="button secondary">Save Draft</button><button id="approve" class="button primary">'+(r.estimate_email_sent_at?'Approved & Emailed':'Approve & Email Estimate')+'</button></div>'+emailStatus+customerStatus+
 '<p class="fine-print">Approving sends the estimate to the customer by email with a PDF attachment and secure response link. Customer approval is recorded separately; scheduling and any additional work are confirmed by the shop.</p></div>';
 $('#status-select').addEventListener('change',e=>save({status:e.target.value}));
 $('#add-line').addEventListener('click',()=>{const box=$('#line-items');box.insertAdjacentHTML('beforeend',itemRow({kind:'Labor',description:'',quantity:1,unitCost:0}));wireRows();calc();});
 $('#save-draft').addEventListener('click',()=>saveEstimate('Estimate Draft'));
 $('#approve').addEventListener('click',()=>{if(r.estimate_email_sent_at){alert('This estimate has already been emailed.');return;}if(!r.email){alert('Add a customer email address before approving and emailing this estimate.');return;}if(!confirm('Approve this estimate and email it to '+r.email+'?'))return;saveEstimate('Approved');});
 $('#parts-type').value=estimate.partsType||'OEM';$('#customer-message').value=estimate.customerMessage||'';
 wireRows();calc();
}
function itemRow(x={}){return '<div class="line-row"><select class="kind"><option '+(x.kind==='Labor'?'selected':'')+'>Labor</option><option '+(x.kind==='Part'?'selected':'')+'>Part</option></select><input class="description" placeholder="Description" value="'+esc(x.description||'')+'"><input class="quantity" type="number" min="0" step="0.1" value="'+Number(x.quantity??1)+'" aria-label="Quantity"><input class="unit-cost" type="number" min="0" step="0.01" value="'+Number(x.unitCost||0)+'" aria-label="Unit cost"><button class="remove-line" type="button" aria-label="Remove line">×</button></div>';}
function wireRows(){document.querySelectorAll('.line-row').forEach(row=>{row.querySelectorAll('input,select').forEach(el=>el.addEventListener('input',calc));row.querySelector('.remove-line').addEventListener('click',()=>{row.remove();calc();});});}
function readItems(){return [...document.querySelectorAll('.line-row')].map(row=>({kind:row.querySelector('.kind').value,description:row.querySelector('.description').value,quantity:Number(row.querySelector('.quantity').value)||0,unitCost:Number(row.querySelector('.unit-cost').value)||0}));}
function totals(items){let labor=0,parts=0;items.forEach(i=>{if(i.kind==='Labor')labor+=i.quantity*135;else if($('#parts-type').value==='Customer Supplied')parts+=i.quantity*i.unitCost;else parts+=i.quantity*i.unitCost*1.4;});return {labor,parts,total:labor+parts};}
function calc(){const t=totals(readItems());$('#totals').innerHTML='<div><span>Labor</span><strong>'+money(t.labor)+'</strong></div><div><span>Parts</span><strong>'+money(t.parts)+'</strong></div><div class="grand-total"><span>Estimated Total</span><strong>'+money(t.total)+'</strong></div>';}
async function saveEstimate(status){const items=readItems();const t=totals(items);await save({status,estimate:{partsType:$('#parts-type').value,customerMessage:$('#customer-message').value,items,totals:t},ownerNotes:$('#owner-notes').value});}
async function save(data){try{const result=await api('/api/requests/'+selected.id,{method:'PATCH',body:JSON.stringify(data)});await loadRequests();if(data.status==='Approved'){if(result.emailSent)alert('Estimate approved and emailed to the customer.');else if(result.emailError)alert('Estimate approved, but the email was not sent: '+result.emailError);}else alert('Request updated.');}catch(e){alert(e.message);}}
function esc(v){return String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
function slug(v){return String(v||'').toLowerCase().replaceAll(' ','-');}
checkLogin();