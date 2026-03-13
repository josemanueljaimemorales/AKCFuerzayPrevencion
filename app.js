
let data=[]
let currentDay=null
let currentType=null

async function loadExcel(){

const response=await fetch('AKC.xlsx')
const arrayBuffer=await response.arrayBuffer()

const workbook=XLSX.read(arrayBuffer)
const sheet=workbook.Sheets[workbook.SheetNames[0]]
data=XLSX.utils.sheet_to_json(sheet)

createApparatusButtons()
createOrientacion()
}

function showSection(type){

currentType=type
hideAll()

if(type==='fuerza')document.getElementById('menu-fuerza').classList.remove('hidden')
if(type==='preventivos')document.getElementById('menu-preventivos').classList.remove('hidden')
if(type==='drills')document.getElementById('menu-drills').classList.remove('hidden')
if(type==='orientacion')document.getElementById('menu-orientacion').classList.remove('hidden')

}

function loadDay(day){

currentDay=day
document.getElementById('routineTitle').innerText=day

hideAll()
document.getElementById('routine').classList.remove('hidden')

renderRoutine()
}

function convertYoutube(url){

if(!url) return ""

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

function renderRoutine(){

const week=document.getElementById('weekSelector').value
const container=document.getElementById('exerciseList')
container.innerHTML=""

const filtered=data.filter(r=>{

if(currentType==='fuerza') return r.Tipo==='Fuerza' && r.Dia===currentDay && String(r.Semana)===week
if(currentType==='preventivos') return r.Tipo==='Preventivo' && r.Dia===currentDay && String(r.Semana)===week

return false

})

filtered.forEach(r=>{

const div=document.createElement("div")

const video=convertYoutube(r.Link)

div.innerHTML=`
<h3>${r.Ejercicio}</h3>
<p>${r.Series||''} x ${r.Reps||''} | Peso: ${r.Peso||''}</p>
<iframe src="${video}" allowfullscreen></iframe>
`

container.appendChild(div)

})

}

function createApparatusButtons(){

const container=document.getElementById('apparatusButtons')
container.innerHTML=""

const apparatus=[...new Set(data.filter(r=>r.Tipo==='Drill').map(r=>r.Aparato))]

apparatus.forEach(a=>{

const btn=document.createElement("button")
btn.innerText=a

btn.onclick=()=>showDrills(a)

container.appendChild(btn)

})

}

function showDrills(app){

hideAll()

document.getElementById('routine').classList.remove('hidden')
document.getElementById('routineTitle').innerText=app+" - Drills"

const container=document.getElementById('exerciseList')
container.innerHTML=""

const drills=data.filter(r=>r.Tipo==='Drill' && r.Aparato===app)

drills.forEach(r=>{

const video=convertYoutube(r.Link)

const div=document.createElement("div")

div.innerHTML=`
<h3>${r.Ejercicio}</h3>
<iframe src="${video}" allowfullscreen></iframe>
`

container.appendChild(div)

})

}

function createOrientacion(){

const container=document.getElementById('orientacionList')
container.innerHTML=""

const orient=data.filter(r=>r.Tipo==='Orientacion')

orient.forEach(r=>{

const video=convertYoutube(r.Link)

const div=document.createElement("div")

div.innerHTML=`
<h3>${r.Ejercicio}</h3>
<iframe src="${video}" allowfullscreen></iframe>
`

container.appendChild(div)

})

}

function hideAll(){

document.getElementById('home').classList.add('hidden')
document.getElementById('menu-fuerza').classList.add('hidden')
document.getElementById('menu-preventivos').classList.add('hidden')
document.getElementById('menu-drills').classList.add('hidden')
document.getElementById('menu-orientacion').classList.add('hidden')
document.getElementById('routine').classList.add('hidden')

}

function goHome(){

hideAll()
document.getElementById('home').classList.remove('hidden')

}

loadExcel()
