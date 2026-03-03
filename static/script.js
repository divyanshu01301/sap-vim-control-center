function updateCounts() {
    fetch("/api/status")
        .then(res => res.json())
        .then(data => {

            document.getElementById("incoming-count").innerText = data.incoming;
            document.getElementById("rejected-count").innerText = data.rejected;

            const ctx = document.getElementById("docChart").getContext("2d");

            new Chart(ctx, {
                type: 'line',
                data: {
                    labels: ['Mon','Tue','Wed','Thu','Fri','Sat','Sun'],
                    datasets: [
                        {
                            label: 'Incoming',
                            data: [2,3,4,data.incoming,3,2,4],
                            borderColor: '#00f5a0',
                            backgroundColor: 'rgba(0,245,160,0.2)',
                            fill: true,
                            tension: 0.4
                        },
                        {
                            label: 'Rejected',
                            data: [1,2,1,data.rejected,2,1,1],
                            borderColor: '#ff3c5f',
                            backgroundColor: 'rgba(255,60,95,0.2)',
                            fill: true,
                            tension: 0.4
                        }
                    ]
                },
                options: {
                    plugins: {
                        legend: { labels: { color: '#ffffff' } }
                    },
                    scales: {
                        x: { ticks: { color: '#aaa' } },
                        y: { ticks: { color: '#aaa' } }
                    }
                }
            });
        });
}

function updateLogs() {
    fetch("/api/logs")
        .then(res => res.json())
        .then(data => {
            document.getElementById("logs-box").innerHTML =
                data.logs.join("<br>");
        });
}

updateCounts();
updateLogs();
setInterval(updateLogs, 4000);