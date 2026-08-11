const S={daily:[],workout:[],cardio:[],settings:{}};const C={};
const n=v=>v===null||v===""||v===undefined?null:Number(v);
const avg=a=>{a=a.filter(Number.isFinite);return a.length?a.reduce((x,y)=>x+y,0)/a.length:null};
const d=v=>{if(!v)return"";if(typeof v==="number"){const x=XLSX.SSF.parse_date_code(v);return`${x.y}-${String(x.m).padStart(2,"0")}-${String(x.d).padStart(2,"0")}`}return String(v).slice(0,10)};
const el=id=>document.getElementById(id);

function kill(k){if(C[k])C[k].destroy();}
function line(k,id,labels,data,label,color){kill(k);C[k]=new Chart(el(id),{type:"line",data:{labels,datasets:[{label,data,borderColor:color,backgroundColor:color+"22",fill:true,tension:.3,spanGaps:true,pointRadius:4}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:"#9aa6c5"}}},scales:{x:{ticks:{color:"#7f8baa"},grid:{color:"#1f2a46"}},y:{ticks:{color:"#7f8baa"},grid:{color:"#1f2a46"}}}}});}
function bar(k,id,labels,data,label,color){kill(k);C[k]=new Chart(el(id),{type:"bar",data:{labels,datasets:[{label,data,backgroundColor:color,borderRadius:7}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:"#9aa6c5"}}},scales:{x:{ticks:{color:"#7f8baa"},grid:{display:false}},y:{beginAtZero:true,ticks:{color:"#7f8baa"},grid:{color:"#1f2a46"}}}}});}

async function load(){
  try{
    const response=await fetch("./fitness_tracker.xlsx?ts="+Date.now(),{cache:"no-store"});
    if(!response.ok) throw new Error(`Excel fetch failed: ${response.status}`);
    const buffer=await response.arrayBuffer();
    const wb=XLSX.read(buffer,{type:"array"});
    const read=name=>XLSX.utils.sheet_to_json(wb.Sheets[name],{defval:null});
    S.daily=read("Daily_Log");
    S.workout=read("Workout_Log");
    S.cardio=read("Cardio_Log");
    S.settings=Object.fromEntries(read("Settings").map(r=>[r.Setting,r.Value]));
    el("status").textContent="Excel connected";
    el("status").style.color="#42e59a";
    render();
  }catch(err){
    console.error(err);
    el("status").textContent="Excel load failed";
    el("status").style.color="#ff6c8d";
  }
}

function render(){
  const start=n(S.settings.Starting_Weight_lb)||245;
  const target=n(S.settings.Target_Weight_lb)||180;
  const weights=S.daily.map(r=>({date:d(r.Date),v:n(r.Body_Weight_lb)})).filter(x=>x.v!==null);
  const current=weights.length?weights[weights.length-1].v:start;
  const lost=start-current;
  const progress=Math.max(0,Math.min(100,lost/(start-target)*100));
  el("title").textContent=S.settings.Dashboard_Title||"Fitness Journey";
  el("currentWeight").textContent=current.toFixed(1);
  el("targetWeight").textContent=target;
  el("totalLost").textContent=lost.toFixed(1);
  el("workoutDays").textContent=S.daily.filter(r=>r.Workout_Day&&String(r.Rest_Day).toLowerCase()!=="yes").length;
  el("cardioSessions").textContent=S.cardio.filter(r=>r.Type).length;
  el("progressPct").textContent=progress.toFixed(0)+"%";
  el("ring").style.setProperty("--p",progress);
  line("weight","weightChart",weights.map(x=>x.date),weights.map(x=>x.v),"Weight (lb)","#20d7ff");
  renderWorkout(); renderCardio(); renderNutrition(); renderRecovery();
}

function renderWorkout(){
  const f=el("muscleFilter");
  const draw=()=>{
    const g=f.value;
    const rows=S.workout.filter(r=>r.Muscle_Group===g);
    el("strengthTitle").textContent=g+" — Load Trend";
    el("volumeTitle").textContent=g+" — Training Volume";
    el("workoutRawTitle").textContent=g+" Raw Workout Data";
    const weighted=rows.filter(r=>n(r.Weight_lb)!==null);
    line("strength","strengthChart",weighted.map(r=>`${d(r.Date)} · ${r.Exercise}`),weighted.map(r=>n(r.Weight_lb)),"Weight (lb)","#9b7bff");
    bar("volume","volumeChart",rows.map(r=>`${d(r.Date)} · ${r.Exercise}`),rows.map(r=>(n(r.Weight_lb)||0)*(n(r.Sets)||0)*(n(r.Reps)||0)),"Volume","#20d7ff");
    el("workoutTable").innerHTML="<thead><tr><th>Date</th><th>Day</th><th>Week</th><th>Muscle Group</th><th>Exercise</th><th>Weight</th><th>Sets</th><th>Reps</th></tr></thead><tbody>"+
      (rows.length?rows.map(r=>`<tr><td>${d(r.Date)}</td><td>${r.Day??""}</td><td>${r.Week??""}</td><td>${r.Muscle_Group}</td><td>${r.Exercise}</td><td>${r.Weight_lb??"Bodyweight"}</td><td>${r.Sets??""}</td><td>${r.Reps??""}</td></tr>`).join(""):`<tr><td class="empty" colspan="8">No ${g} data logged yet.</td></tr>`)+"</tbody>";
  };
  f.onchange=draw; draw();
}

function renderCardio(){
  const rows=S.cardio.filter(r=>r.Type);
  el("cardioTotal").textContent=rows.length;
  el("cardioMinutes").textContent=rows.reduce((s,r)=>s+(n(r.Duration_min)||0),0);
  el("cardioCalories").textContent=rows.reduce((s,r)=>s+(n(r.Calories)||0),0);
  const mix={}; rows.forEach(r=>mix[r.Type]=(mix[r.Type]||0)+1);
  bar("cardioMix","cardioMixChart",Object.keys(mix),Object.values(mix),"Sessions","#20d7ff");
  line("cardioDuration","cardioDurationChart",rows.map(r=>d(r.Date)),rows.map(r=>n(r.Duration_min)),"Minutes","#42e59a");
  el("cardioTable").innerHTML="<thead><tr><th>Date</th><th>Day</th><th>Week</th><th>Type</th><th>Duration</th><th>Calories</th><th>Distance</th><th>Details</th></tr></thead><tbody>"+
  rows.map(r=>`<tr><td>${d(r.Date)}</td><td>${r.Day??""}</td><td>${r.Week??""}</td><td>${r.Type}</td><td>${r.Duration_min??""}</td><td>${r.Calories??""}</td><td>${r.Distance??""}</td><td>${r.Details??""}</td></tr>`).join("")+"</tbody>";
}

function renderNutrition(){
  const rows=S.daily.filter(r=>n(r.Calories)!==null||n(r.Protein_g)!==null);
  const ca=avg(rows.map(r=>n(r.Calories))), pa=avg(rows.map(r=>n(r.Protein_g)));
  el("avgCalories").textContent=ca?Math.round(ca):"—";
  el("avgProtein").textContent=pa?Math.round(pa):"—";
  el("proteinTarget").textContent=n(S.settings.Protein_Target_g)||160;
  bar("calories","calorieChart",rows.map(r=>"D"+r.Day),rows.map(r=>n(r.Calories)),"Calories","#ffb14a");
  line("protein","proteinChart",rows.map(r=>"D"+r.Day),rows.map(r=>n(r.Protein_g)),"Protein (g)","#9b7bff");
  el("nutritionTable").innerHTML="<thead><tr><th>Date</th><th>Day</th><th>Week</th><th>Calories</th><th>Protein</th><th>Workout / Day</th><th>Notes</th></tr></thead><tbody>"+
  rows.map(r=>`<tr><td>${d(r.Date)}</td><td>${r.Day??""}</td><td>${r.Week??""}</td><td>${r.Calories??""}</td><td>${r.Protein_g??""}</td><td>${r.Workout_Day??""}</td><td>${r.Notes??""}</td></tr>`).join("")+"</tbody>";
}

function renderRecovery(){
  const rest=S.daily.filter(r=>String(r.Rest_Day).toLowerCase()==="yes");
  const energies=S.daily.map(r=>n(r.Energy_1_to_10)).filter(Number.isFinite);
  el("restDays").textContent=rest.length;
  el("avgEnergy").textContent=energies.length?(energies.reduce((a,b)=>a+b,0)/energies.length).toFixed(1):"—";
  const rows=S.daily.filter(r=>String(r.Rest_Day).toLowerCase()==="yes"||n(r.Energy_1_to_10)!==null);
  el("recoveryTable").innerHTML="<thead><tr><th>Date</th><th>Day</th><th>Week</th><th>Rest Day</th><th>Energy /10</th><th>Notes</th></tr></thead><tbody>"+
  (rows.length?rows.map(r=>`<tr><td>${d(r.Date)}</td><td>${r.Day??""}</td><td>${r.Week??""}</td><td>${r.Rest_Day}</td><td>${r.Energy_1_to_10??""}</td><td>${r.Notes??""}</td></tr>`).join(""):`<tr><td class="empty" colspan="6">No recovery data logged yet.</td></tr>`)+"</tbody>";
}

document.querySelectorAll("nav button").forEach(btn=>btn.onclick=()=>{
  document.querySelectorAll("nav button").forEach(x=>x.classList.remove("active"));
  document.querySelectorAll(".tab").forEach(x=>x.classList.remove("active"));
  btn.classList.add("active");
  el(btn.dataset.tab).classList.add("active");
  setTimeout(()=>Object.values(C).forEach(c=>c.resize()),30);
});

load();
