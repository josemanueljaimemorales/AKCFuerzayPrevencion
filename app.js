
let data = [];

async function loadExcel(){
    const response = await fetch('AKC_MASTER_Training_System.xlsx');
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer);
    const sheet = workbook.Sheets['Base_Datos'];
    data = XLSX.utils.sheet_to_json(sheet);
}

function openSection(type){

    const container = document.getElementById('content');
    container.innerHTML = "";

    let filtered = data.filter(r => r.Tipo === type);

    filtered.forEach(e => {

        const div = document.createElement("div");
        div.className = "exercise";

        div.innerHTML = `
        <h3>${e.Ejercicio || ""}</h3>
        <p>${e.Descripcion || ""}</p>
        <p>${e.Series || ""} x ${e.Reps || ""}</p>
        <a href="${e.Link || "#"}" target="_blank">Ver video</a>
        `;

        container.appendChild(div);
    });

}

loadExcel();
