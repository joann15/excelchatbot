const API_URL = "https://excelchatbot.onrender.com";
let dashboardData = null;

// UPLOAD 
function uploadFiles() {

    const files = document.getElementById("files").files;

let formData = new FormData();

for (let file of files) {
    formData.append("files", file);
}
    document.getElementById("status").innerText = "Processing...";

    fetch(`${API_URL}/upload`, {
        method: "POST",
        body: formData
    })
    .then(res => res.json())
    .then(data => {

        document.getElementById("status").innerText =
            "Uploaded " + data.files + " files";

        loadDashboard();  
    })
    .catch(err => {
        console.log(err);
        document.getElementById("status").innerText = "Error uploading";
    });
}


// ================= DASHBOARD =================
function loadDashboard() {

    fetch(`${API_URL}/dashboard`)
        .then(res => res.json())
        .then(data => {

            document.getElementById("dashboard").style.display = "block";

            document.getElementById("totalTasks").innerText =
                data.total_tasks;

            document.getElementById("doneCount").innerText =
                data.status_breakdown["Done"] || 0;
            
            document.getElementById("dueCount").innerText =
                data.status_breakdown["Due"] || 0;
                
                const done = data.status_breakdown["Done"] || 0;
const due = data.status_breakdown["Due"] || 0;
const half = data.status_breakdown["Half-Done"] || 0;

const totalAssignments =
Object.values(data.status_breakdown).reduce((a, b) => a + b, 0);

const donePercent =
totalAssignments > 0 ? (done / totalAssignments) * 100 : 0;

const duePercent =
totalAssignments > 0 ? (due / totalAssignments) * 100 : 0;

const halfPercent =
totalAssignments > 0 ? (half / totalAssignments) * 100 : 0;

// Show completion percentage (Done only)

document.getElementById("completionRate").innerText =
    donePercent.toFixed(1) + "%";

// Update the 3 colored sections
document.getElementById("doneBar").style.width = donePercent + "%";
document.getElementById("dueBar").style.width = duePercent + "%";
document.getElementById("halfBar").style.width = halfPercent + "%";
                
        
            // destroy old chart if user uploads again
            if (window.workloadChart && typeof window.workloadChart.destroy === "function") {
                window.workloadChart.destroy();}
const employees = Object.keys(data.employee_status);

// build datasets dynamically
const statuses = Object.keys(data.status_breakdown);

const colors = data.status_colors;

const datasets = statuses.map(status => ({
    label: status,
    data: employees.map(emp =>
        data.employee_status[emp]?.[status] || 0
    ),
    backgroundColor: colors[status]
}));

window.workloadChart = new Chart(
    document.getElementById("employeeChart"),
    {
        type: "bar",
        data: {
            labels: employees,
            datasets: datasets
        },
        options: {
            responsive: true,
            scales: {
                x: {
                    stacked: true
                },
                y: {
                    stacked: true
                }
            }
        }
    }
);
            
            
        });
}
           

// ================= CHAT =================
function sendMessage() {

    const msg = document.getElementById("msg").value;

    const responseBox = document.getElementById("response");

    responseBox.style.display = "block";
    responseBox.innerHTML = "Thinking...";

    fetch(`${API_URL}/chat`, {
        method: "POST",
        headers: {
            "Content-Type": "application/json"
        },
        body: JSON.stringify({ message: msg })
    })
    .then(res => res.json())
    .then(data => {

    responseBox.innerText = data.answer;

    // Refresh dashboard after task creation
    setTimeout(loadDashboard, 1500);

})
    .catch(err => {
        responseBox.innerText = "Error getting response";
    });
}