let designId = null;
let connected = false;

const $ = (id)=>document.getElementById(id);
const setStatus = (s)=> $("status").textContent = "Status: " + s;

function extractId(u) {
  const m = (u||'').match(/canva\.com\/design\/([A-Za-z0-9_\-]+)/);
  return m?.[1] || null;
}

// Restore simple connected state if sess cookie exists
function hasSessCookie(){
  return document.cookie.split(/;\s*/).some(x=>x.startsWith('sess='));
}
function refreshButtons(){
  $("export").disabled = !designId || !connected;
  $("resize").disabled = !designId || !connected;
}

$("resolve").onclick = ()=>{
  const id = extractId($("link").value.trim());
  designId = id;
  $("design").textContent = id ? ("Design ID: " + id) : "Invalid link";
  setStatus(id ? "Ready" : "Invalid link");
  refreshButtons();
};

$("connect").onclick = ()=>{
  window.location.href = "/auth/canva/start";
};

$("disconnect").onclick = async ()=>{
  await fetch("/api/disconnect", {method:"POST"});
  connected = false;
  setStatus("Disconnected");
  refreshButtons();
};

$("export").onclick = async ()=>{
  if(!designId || !connected) return;
  setStatus("Exporting original PPTX...");
  const r = await fetch("/api/export", {
    method: "POST", headers: {"Content-Type":"application/json"},
    body: JSON.stringify({ designId })
  });
  const j = await r.json();
  if (r.ok) setStatus("Original PPTX ready: " + j.s3Url);
  else setStatus("Error: " + (j.error||"export_failed"));
};

$("resize").onclick = async ()=>{
  if(!designId || !connected) return;
  setStatus("Creating 4:3 copy & exporting PPTX...");
  const r = await fetch("/api/resizeExport", {
    method: "POST", headers: {"Content-Type":"application/json"},
    body: JSON.stringify({ designId, width:1600, height:1200 })
  });
  const j = await r.json();
  if (r.ok) setStatus("4:3 PPTX ready: " + j.s3Url);
  else setStatus("Error: " + (j.error||"resize_export_failed"));
};

// Init
connected = hasSessCookie();
refreshButtons();
if (connected) setStatus("Connected. Paste a link and resolve the Design ID.");