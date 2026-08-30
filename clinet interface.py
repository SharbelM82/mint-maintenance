<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Mint Maintenance LLC - Request Service</title>
    <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="bg-slate-50 font-sans text-slate-800">
    <header class="bg-white border-b border-slate-200 py-4 px-8 flex justify-between items-center">
        <h1 class="text-xl font-bold text-emerald-700 tracking-wide">MINT MAINTENANCE LLC</h1>
        <a href="/admin" class="text-sm font-semibold text-slate-600 hover:text-emerald-700">Admin Dashboard</a>
    </header>

    <main class="max-w-xl mx-auto my-12 p-8 bg-white rounded-xl shadow-md border border-slate-100">
        <h2 class="text-2xl font-bold mb-2 text-slate-900">Request Facility Service</h2>
        <p class="text-sm text-slate-500 mb-6">Select a trade service to get an instant cost quotation.</p>

        <form id="bookingForm" onsubmit="submitForm(event)" class="space-y-4">
            <div>
                <label class="block text-sm font-semibold mb-1">Service Category</label>
                <select id="category" required class="w-full p-2.5 border border-slate-300 rounded-lg">
                    <option value="HVAC">HVAC & Heating ($120 base)</option>
                    <option value="Plumbing">Plumbing ($95 base)</option>
                    <option value="Electrical">Electrical ($110 base)</option>
                    <option value="Handyman">Handyman Services ($75 base)</option>
                    <option value="Renovation">Full Renovation ($250 base)</option>
                </select>
            </div>

            <div class="grid grid-cols-2 gap-4">
                <div>
                    <label class="block text-sm font-semibold mb-1">Full Name</label>
                    <input type="text" id="clientName" required placeholder="John Doe" class="w-full p-2.5 border border-slate-300 rounded-lg">
                </div>
                <div>
                    <label class="block text-sm font-semibold mb-1">Email</label>
                    <input type="email" id="clientEmail" required placeholder="john@example.com" class="w-full p-2.5 border border-slate-300 rounded-lg">
                </div>
            </div>

            <div>
                <label class="block text-sm font-semibold mb-1">Service Address</label>
                <input type="text" id="address" required placeholder="123 Business Rd" class="w-full p-2.5 border border-slate-300 rounded-lg">
            </div>

            <div class="grid grid-cols-2 gap-4">
                <div>
                    <label class="block text-sm font-semibold mb-1">Zip Code</label>
                    <input type="text" id="zipCode" required placeholder="90210" class="w-full p-2.5 border border-slate-300 rounded-lg">
                </div>
                <div>
                    <label class="block text-sm font-semibold mb-1">Urgency</label>
                    <select id="urgency" class="w-full p-2.5 border border-slate-300 rounded-lg">
                        <option value="STANDARD">Standard Service</option>
                        <option value="SAME_DAY">Same-Day (+25%)</option>
                        <option value="EMERGENCY_24_7">24/7 Emergency (+50%)</option>
                    </select>
                </div>
            </div>

            <div>
                <label class="block text-sm font-semibold mb-1">Issue Description</label>
                <textarea id="description" rows="3" required placeholder="Describe the issue..." class="w-full p-2.5 border border-slate-300 rounded-lg"></textarea>
            </div>

            <button type="submit" class="w-full py-3 bg-emerald-700 text-white font-bold rounded-lg hover:bg-emerald-800 transition">
                Submit & Request Quote
            </button>
        </form>

        <div id="resultModal" class="hidden mt-6 p-4 bg-emerald-50 border border-emerald-200 rounded-lg text-center">
            <h3 class="text-lg font-bold text-emerald-900">Service Request Created!</h3>
            <p class="text-2xl font-black text-emerald-700 my-2" id="costResult">$0.00</p>
            <p class="text-xs text-slate-600">Status: <span id="statusResult" class="font-bold">PENDING</span></p>
        </div>
    </main>

    <script>
        async function submitForm(e) {
            e.preventDefault();
            const payload = {
                category: document.getElementById('category').value,
                clientName: document.getElementById('clientName').value,
                clientEmail: document.getElementById('clientEmail').value,
                address: document.getElementById('address').value,
                zipCode: document.getElementById('zipCode').value,
                urgency: document.getElementById('urgency').value,
                description: document.getElementById('description').value
            };

            const res = await fetch('/api/tickets', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });

            const data = await res.json();
            if (data.success) {
                document.getElementById('resultModal').classList.remove('hidden');
                document.getElementById('costResult').innerText = `$${data.estimated_cost.toFixed(2)}`;
                document.getElementById('statusResult').innerText = data.status;
            }
        }
    </script>
</body>
</html>