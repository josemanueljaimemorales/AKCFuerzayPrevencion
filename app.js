
let data=[]
let baseDate=new Date('2025-03-17')

async function loadExcel(){

const response=await fetch('AKC.xlsx')
const buffer=await response.arrayBuffer()

const workbook=XLSX.read(buffer)
const sheet=workbook.Sheets[workbook.SheetNames[0]]
data=XLSX.utils.sheet_to_json(sheet)

console.log("Excel cargado",data.length)

}

function getCycleWeek(){

const now=new Date()
const diff=Math.floor((now-baseDate)/(1000*60*60*24*7))

const pattern=[1,2,3,2]

return pattern[diff%pattern.length]

}

function openSection(type){

hideAll()

if(type==="Fuerza"||type==="Preventivo"){

const days=type==="Fuerza"?["Lunes","Miercoles","Viernes"]:["Jueves"]

const menu=document.getElementById("menu")
menu.innerHTML=""

days.forEach(d=>{

const btn=document.createElement("button")
btn.innerText=d
btn.onclick=()=>showRoutine(type,d)

menu.appendChild(btn)

})

menu.classList.remove("hidden")

}

else if(type==="Drill"||type==="F ESP APA"){

const menu=document.getElementById("menu")
menu.innerHTML=""

const aparatos=[...new Set(data.filter(r=>r.Tipo===type).map(r=>r.Aparato))]

aparatos.forEach(a=>{

const btn=document.createElement("button")
btn.innerText=a
btn.onclick=()=>showApparatus(type,a)

menu.appendChild(btn)

})

menu.classList.remove("hidden")

}

else if(type==="Orientacion"){

showOrientation()

}

}

function convert(url){

if(!url)return ""

if(url.includes("shorts")){
const id=url.split("/shorts/")[1].split("?")[0]
return "https://www.youtube.com/embed/"+id
}

if(url.includes("watch?v=")){
const id=url.split("v=")[1]
return "https://www.youtube.com/embed/"+id
}

return url

}

function showRoutine(type,day){

hideAll()

const week=getCycleWeek()

const rows=data.filter(r=>r.Tipo===type&&r.Dia===day&&String(r.Semana)===String(week))

render(rows,day+" - Semana "+week)

}

function showApparatus(type,app){

hideAll()

const rows=data.filter(r=>r.Tipo===type&&r.Aparato===app)

render(rows,app)

}

function showOrientation(){

hideAll()

const rows=data.filter(r=>r.Tipo==="Orientacion")

render(rows,"Orientación")

}

function render(rows,title){

document.getElementById("title").innerText=title

const container=document.getElementById("list")
container.innerHTML=""

rows.forEach(r=>{

const video=convert(r.Link||"")

const div=document.createElement("div")

div.innerHTML=`
<h3>${r.Ejercicio||""}</h3>
<p>${r.Series||""} x ${r.Reps||""} | Peso: ${r.Peso||""}</p>
<iframe src="${video}" allowfullscreen></iframe>
`

container.appendChild(div)

})

document.getElementById("routine").classList.remove("hidden")

}

function hideAll(){

document.getElementById("menu").classList.add("hidden")
document.getElementById("routine").classList.add("hidden")

}

function goHome(){

hideAll()

}

window.onload=loadExcel
