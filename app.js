const S={daily:[],workout:[],cardio:[],settings:{}};const C={};
const n=v=>v===null||v===""||v===undefined?null:Number(v);
const avg=a=>{a=a.filter(Number.isFinite);return a.length?a.reduce((x,y)=>x+y,0)/a.length:null};
const d=v=>{
  if(v===null||v===undefined||v==="") return "";
  if(typeof v==="number"){
    const x=XLSX.SSF.parse_date_code(v);
    if(x) return `${x.y}-${String(x.m).padStart(2,"0")}-${String(x.d).padStart(2,"0")}`;
  }
  if(v instanceof Date) return v.toISOString().slice(0,10);
  return String(v).slice(0,10);
};
const el=id=>document.getElementById(id);
const txt=v=>v===null||v===undefined||v===""?"—":String(v);

function rowsByColumns(sheet, columns){
  const grid=XLSX.utils.sheet_to_json(sheet,{header:1,defval:null,raw:true});
  if(!grid.length) return [];
  const headers=(grid[0]||[]).map(h=>String(h??"").trim());
  const idx={};
  columns.forEach(c=>idx[c]=headers.indexOf(c));
  return grid.slice(1).map(row=>{
    const o={};
    columns.forEach(c=>o[c]=idx[c]>=0 ? (row[idx[c]] ?? null) : null);
    return o;
  });
}

function kill(k){ if(C[k]) C[k].destroy(); }
function line(k,id,labels,data,label,color){
  kill(k);
  const canvas=el(id);
  if(!canvas) return;
  C[k]=new Chart(canvas,{
    type:"line",
    data:{labels,datasets:[{label,data,borderColor:color,backgroundColor:color+"22",fill:true,tension:.3,spanGaps:true,pointRadius:4}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:"#9aa6c5"}}},
      scales:{x:{ticks:{color:"#7f8baa"},grid:{color:"#1f2a46"}},y:{ticks:{color:"#7f8baa"},grid:{color:"#1f2a46"}}}}
  });
}
function bar(k,id,labels,data,label,color){
  kill(k);
  const canvas=el(id);
  if(!canvas) return;
  C[k]=new Chart(canvas,{
    type:"bar",
    data:{labels,datasets:[{label,data,backgroundColor:color,borderRadius:7}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:"#9aa6c5"}}},
      scales:{x:{ticks:{color:"#7f8baa"},grid:{display:false}},y:{beginAtZero:true,ticks:{color:"#7f8baa"},grid:{color:"#1f2a46"}}}}
  });
}

async function load(){
  try{
    const response=await fetch("./fitness_tracker.xlsx?ts="+Date.now(),{cache:"no-store"});
    if(!response.ok) throw new Error(`Excel fetch failed: ${response.status}`);
    const buffer=await response.arrayBuffer();
    const wb=XLSX.read(buffer,{type:"array"});

    const dailySheet=wb.Sheets["Daily_Log"];
    const workoutSheet=wb.Sheets["Workout_Log"];
    const cardioSheet=wb.Sheets["Cardio_Log"];
    const settingsSheet=wb.Sheets["Settings"];

    if(!dailySheet) throw new Error("Daily_Log sheet not found");
    if(!workoutSheet) throw new Error("Workout_Log sheet not found");
    if(!cardioSheet) throw new Error("Cardio_Log sheet not found");
    if(!settingsSheet) throw new Error("Settings sheet not found");

    S.daily=rowsByColumns(dailySheet,[
      "Date","Day","Week","Workout_Day","Calories","Protein_g",
      "Body_Weight_lb","Rest_Day","Energy_1_to_10","Notes"
    ]).filter(r=>n(r.Day)!==null);

    S.workout=rowsByColumns(workoutSheet,[
      "Date","Day","Week","Muscle_Group","Exercise","Weight_lb","Sets","Reps"
    ]).filter(r=>n(r.Day)!==null && r.Exercise);

    S.cardio=rowsByColumns(cardioSheet,[
      "Date","Day","Week","Type","Duration_min","Calories","Distance","Details"
    ]).filter(r=>n(r.Day)!==null && r.Type);

    const settingsRows=rowsByColumns(settingsSheet,["Setting","Value"]).filter(r=>r.Setting);
    S.settings=Object.fromEntries(settingsRows.map(r=>[String(r.Setting).trim(),r.Value]));

    const status=el("status");
    if(status){
      status.textContent=`Excel connected • ${S.daily.length} days`;
      status.style.color="#42e59a";
    }
    console.log("Loaded rows", {daily:S.daily.length, workout:S.workout.length, cardio:S.cardio.length});
    render();
  }catch(err){
    console.error("Dashboard load error:",err);
    const status=el("status");
    if(status){
      status.textContent="Excel load failed";
      status.style.color="#ff6c8d";
    }
  }
}

function render(){
  const start=n(S.settings.Starting_Weight_lb)||245;
  const target=n(S.settings.Target_Weight_lb)||180;
  const weights=S.daily.map(r=>({date:d(r.Date),v:n(r.Body_Weight_lb)})).filter(x=>x.v!==null);
  const current=weights.length?weights[weights.length-1].v:start;
  const lost=start-current;
  const progress=Math.max(0,Math.min(100,lost/(start-target)*100));

  if(el("title")) el("title").textContent=S.settings.Dashboard_Title||"Fitness Journey";
  if(el("currentWeight")) el("currentWeight").textContent=current.toFixed(1);
  if(el("targetWeight")) el("targetWeight").textContent=target;
  if(el("totalLost")) el("totalLost").textContent=lost.toFixed(1);
  if(el("workoutDays")) el("workoutDays").textContent=S.daily.filter(r=>r.Workout_Day&&String(r.Rest_Day).toLowerCase()!=="yes").length;
  if(el("cardioSessions")) el("cardioSessions").textContent=S.cardio.length;
  if(el("progressPct")) el("progressPct").textContent=progress.toFixed(0)+"%";
  if(el("ring")) el("ring").style.setProperty("--p",progress);

  line("weight","weightChart",weights.map(x=>x.date),weights.map(x=>x.v),"Weight (lb)","#20d7ff");

  renderDailyView();
  renderWorkout();
  renderCardio();
  renderNutrition();
  renderRecovery();
  renderDailyRawData();
}

function renderDailyView(){
  const filter=el("dayFilter");
  if(!filter) return;

  const days=[...new Set(S.daily.map(r=>n(r.Day)).filter(Number.isFinite))].sort((a,b)=>a-b);
  filter.innerHTML=days.map(day=>`<option value="${day}">Day ${day}</option>`).join("");

  if(!days.length){
    filter.innerHTML='<option>No days found</option>';
    return;
  }

  filter.value=String(days[days.length-1]);

  const draw=()=>{
    const day=n(filter.value);
    const daily=S.daily.find(r=>n(r.Day)===day)||{};
    const workouts=S.workout.filter(r=>n(r.Day)===day);
    const cardio=S.cardio.filter(r=>n(r.Day)===day);

    if(el("dayDate")) el("dayDate").textContent=txt(d(daily.Date));
    if(el("dayWeek")) el("dayWeek").textContent=txt(daily.Week);
    if(el("dayCalories")) el("dayCalories").textContent=txt(daily.Calories);
    if(el("dayProtein")) el("dayProtein").textContent=txt(daily.Protein_g);
    if(el("dayWeight")) el("dayWeight").textContent=txt(daily.Body_Weight_lb);
    if(el("dayWorkoutSummary")) el("dayWorkoutSummary").textContent=txt(daily.Workout_Day);
    if(el("dayRest")) el("dayRest").textContent=txt(daily.Rest_Day);
    if(el("dayEnergy")) el("dayEnergy").textContent=n(daily.Energy_1_to_10)!==null?`${daily.Energy_1_to_10}/10`:"—";
    if(el("dayNotes")) el("dayNotes").textContent=txt(daily.Notes);

    if(el("dayExerciseCount")) el("dayExerciseCount").textContent=`${workouts.length} exercise${workouts.length===1?"":"s"}`;
    if(el("dailyWorkoutTable")){
      el("dailyWorkoutTable").innerHTML=
        "<thead><tr><th>Exercise</th><th>Muscle Group</th><th>Weight</th><th>Sets</th><th>Reps</th></tr></thead><tbody>"+
        (workouts.length
          ? workouts.map(r=>`<tr><td>${txt(r.Exercise)}</td><td>${txt(r.Muscle_Group)}</td><td>${r.Weight_lb??"Bodyweight"}</td><td>${txt(r.Sets)}</td><td>${txt(r.Reps)}</td></tr>`).join("")
          : `<tr><td class="empty" colspan="5">No detailed workout rows logged for Day ${day}.</td></tr>`)
        +"</tbody>";
    }

    if(el("dayCardioCount")) el("dayCardioCount").textContent=`${cardio.length} entr${cardio.length===1?"y":"ies"}`;
    if(el("dailyCardioTable")){
      el("dailyCardioTable").innerHTML=
        "<thead><tr><th>Type</th><th>Duration</th><th>Calories</th><th>Distance</th><th>Details</th></tr></thead><tbody>"+
        (cardio.length
          ? cardio.map(r=>`<tr><td>${txt(r.Type)}</td><td>${txt(r.Duration_min)}</td><td>${txt(r.Calories)}</td><td>${txt(r.Distance)}</td><td>${txt(r.Details)}</td></tr>`).join("")
          : `<tr><td class="empty" colspan="5">No cardio entry logged for Day ${day}.</td></tr>`)
        +"</tbody>";
    }
  };

  filter.onchange=draw;
  draw();
}

function renderWorkout(){
  const f=el("muscleFilter");
  if(!f) return;
  const draw=()=>{
    const g=f.value;
    const rows=S.workout.filter(r=>r.Muscle_Group===g);
    if(el("strengthTitle")) el("strengthTitle").textContent=g+" — Load Trend";
    if(el("volumeTitle")) el("volumeTitle").textContent=g+" — Training Volume";
    if(el("workoutRawTitle")) el("workoutRawTitle").textContent=g+" Raw Workout Data";

    const weighted=rows.filter(r=>n(r.Weight_lb)!==null);
    line("strength","strengthChart",weighted.map(r=>`${d(r.Date)} · ${r.Exercise}`),weighted.map(r=>n(r.Weight_lb)),"Weight (lb)","#9b7bff");
    bar("volume","volumeChart",rows.map(r=>`${d(r.Date)} · ${r.Exercise}`),rows.map(r=>(n(r.Weight_lb)||0)*(n(r.Sets)||0)*(n(r.Reps)||0)),"Volume","#20d7ff");

    if(el("workoutTable")){
      el("workoutTable").innerHTML="<thead><tr><th>Date</th><th>Day</th><th>Week</th><th>Muscle Group</th><th>Exercise</th><th>Weight</th><th>Sets</th><th>Reps</th></tr></thead><tbody>"+
      (rows.length
       ? rows.map(r=>`<tr><td>${d(r.Date)}</td><td>${txt(r.Day)}</td><td>${txt(r.Week)}</td><td>${txt(r.Muscle_Group)}</td><td>${txt(r.Exercise)}</td><td>${r.Weight_lb??"Bodyweight"}</td><td>${txt(r.Sets)}</td><td>${txt(r.Reps)}</td></tr>`).join("")
       : `<tr><td class="empty" colspan="8">No ${g} data logged yet.</td></tr>`)
      +"</tbody>";
    }
  };
  f.onchange=draw;
  draw();
}

function renderCardio(){
  const rows=S.cardio;
  if(el("cardioTotal")) el("cardioTotal").textContent=rows.length;
  if(el("cardioMinutes")) el("cardioMinutes").textContent=rows.reduce((s,r)=>s+(n(r.Duration_min)||0),0);
  if(el("cardioCalories")) el("cardioCalories").textContent=rows.reduce((s,r)=>s+(n(r.Calories)||0),0);
  const mix={}; rows.forEach(r=>mix[r.Type]=(mix[r.Type]||0)+1);
  bar("cardioMix","cardioMixChart",Object.keys(mix),Object.values(mix),"Sessions","#20d7ff");
  line("cardioDuration","cardioDurationChart",rows.map(r=>d(r.Date)),rows.map(r=>n(r.Duration_min)),"Minutes","#42e59a");
  if(el("cardioTable")){
    el("cardioTable").innerHTML="<thead><tr><th>Date</th><th>Day</th><th>Week</th><th>Type</th><th>Duration</th><th>Calories</th><th>Distance</th><th>Details</th></tr></thead><tbody>"+
    rows.map(r=>`<tr><td>${d(r.Date)}</td><td>${txt(r.Day)}</td><td>${txt(r.Week)}</td><td>${txt(r.Type)}</td><td>${txt(r.Duration_min)}</td><td>${txt(r.Calories)}</td><td>${txt(r.Distance)}</td><td>${txt(r.Details)}</td></tr>`).join("")+
    "</tbody>";
  }
}

function renderNutrition(){
  const rows=S.daily.filter(r=>n(r.Calories)!==null||n(r.Protein_g)!==null);
  const ca=avg(rows.map(r=>n(r.Calories))), pa=avg(rows.map(r=>n(r.Protein_g)));
  if(el("avgCalories")) el("avgCalories").textContent=ca?Math.round(ca):"—";
  if(el("avgProtein")) el("avgProtein").textContent=pa?Math.round(pa):"—";
  if(el("proteinTarget")) el("proteinTarget").textContent=n(S.settings.Protein_Target_g)||160;
  bar("calories","calorieChart",rows.map(r=>"D"+r.Day),rows.map(r=>n(r.Calories)),"Calories","#ffb14a");
  line("protein","proteinChart",rows.map(r=>"D"+r.Day),rows.map(r=>n(r.Protein_g)),"Protein (g)","#9b7bff");
  if(el("nutritionTable")){
    el("nutritionTable").innerHTML="<thead><tr><th>Date</th><th>Day</th><th>Week</th><th>Calories</th><th>Protein</th><th>Workout / Day</th><th>Notes</th></tr></thead><tbody>"+
    (rows.length
     ? rows.map(r=>`<tr><td>${d(r.Date)}</td><td>${txt(r.Day)}</td><td>${txt(r.Week)}</td><td>${txt(r.Calories)}</td><td>${txt(r.Protein_g)}</td><td>${txt(r.Workout_Day)}</td><td>${txt(r.Notes)}</td></tr>`).join("")
     : '<tr><td class="empty" colspan="7">No calorie or protein values logged yet.</td></tr>')+
    "</tbody>";
  }
}

function renderRecovery(){
  const rest=S.daily.filter(r=>String(r.Rest_Day).toLowerCase()==="yes");
  const energies=S.daily.map(r=>n(r.Energy_1_to_10)).filter(Number.isFinite);
  if(el("restDays")) el("restDays").textContent=rest.length;
  if(el("avgEnergy")) el("avgEnergy").textContent=energies.length?(energies.reduce((a,b)=>a+b,0)/energies.length).toFixed(1):"—";
  const rows=S.daily.filter(r=>String(r.Rest_Day).toLowerCase()==="yes"||n(r.Energy_1_to_10)!==null);
  if(el("recoveryTable")){
    el("recoveryTable").innerHTML="<thead><tr><th>Date</th><th>Day</th><th>Week</th><th>Rest Day</th><th>Energy /10</th><th>Notes</th></tr></thead><tbody>"+
    (rows.length
     ? rows.map(r=>`<tr><td>${d(r.Date)}</td><td>${txt(r.Day)}</td><td>${txt(r.Week)}</td><td>${txt(r.Rest_Day)}</td><td>${txt(r.Energy_1_to_10)}</td><td>${txt(r.Notes)}</td></tr>`).join("")
     : '<tr><td class="empty" colspan="6">No recovery data logged yet.</td></tr>')+
    "</tbody>";
  }
}

function renderDailyRawData(){
  const rows=[...S.daily].sort((a,b)=>(n(a.Day)||0)-(n(b.Day)||0));
  if(el("dailyRawCount")) el("dailyRawCount").textContent=`${rows.length} days`;
  if(el("dailyRawTable")){
    el("dailyRawTable").innerHTML="<thead><tr><th>Date</th><th>Day</th><th>Week</th><th>Workout Day</th><th>Calories</th><th>Protein (g)</th><th>Body Weight (lb)</th><th>Rest Day</th><th>Energy /10</th><th>Notes</th></tr></thead><tbody>"+
    rows.map(r=>`<tr><td>${d(r.Date)}</td><td>${txt(r.Day)}</td><td>${txt(r.Week)}</td><td>${txt(r.Workout_Day)}</td><td>${txt(r.Calories)}</td><td>${txt(r.Protein_g)}</td><td>${txt(r.Body_Weight_lb)}</td><td>${txt(r.Rest_Day)}</td><td>${txt(r.Energy_1_to_10)}</td><td>${txt(r.Notes)}</td></tr>`).join("")+
    "</tbody>";
  }
}

document.querySelectorAll("nav button").forEach(btn=>btn.onclick=()=>{
  document.querySelectorAll("nav button").forEach(x=>x.classList.remove("active"));
  document.querySelectorAll(".tab").forEach(x=>x.classList.remove("active"));
  btn.classList.add("active");
  const section=el(btn.dataset.tab);
  if(section) section.classList.add("active");
  setTimeout(()=>Object.values(C).forEach(c=>c.resize()),30);
});

load();
