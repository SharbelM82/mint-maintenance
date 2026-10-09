"""Fill an empty database with sample work orders so the app can be explored.

    python seed_demo.py           (refuses to run if work orders already exist)
"""
import random
from datetime import date, timedelta

from mint import create_app, db

CLIENTS = ["ServiceChannel", "Corrigo", "FMPilot", "Verisae", "Direct client"]
STORES = [("Dollar Plus #1182", "Houston, TX"), ("Fresh Mart #204", "Dallas, TX"),
          ("QuickFuel #77", "Austin, TX"), ("Bright Pharmacy #310", "San Antonio, TX"),
          ("Urban Fitness", "Plano, TX"), ("Coffee Corner #12", "Irving, TX")]
JOBS = ["HVAC - no cooling", "Plumbing - leaking sink", "Electrical - lights out",
        "Door repair", "Roof leak", "Ice machine repair", "Handyman - ceiling tiles"]
PEOPLE = ["Charbel", "Maria", "Tony", "Rami"]
TECHS = ["J. Lopez", "K. Smith", "A. Haddad", "D. Brown"]


def main():
    app = create_app()
    rnd = random.Random(7)
    with app.app_context():
        if db.all_work_orders():
            print("Work orders already exist; nothing seeded.")
            return
        today = date.today()
        for n in range(1, 41):
            store, loc = rnd.choice(STORES)
            date_in = today - timedelta(days=rnd.randint(0, 150))
            cost = round(rnd.uniform(150, 2500), 2)
            markup = rnd.choice([0.3, 0.35, 0.4, 0.5])
            nte = round(cost * (1 + markup) * rnd.choice([0.9, 1.1, 1.2, 1.5]), -1)
            status = rnd.choice(["New", "Scheduled", "In progress", "Awaiting parts",
                                 "Awaiting approval", "Completed", "Completed", "Completed"])
            wo = {"wo_number": f"WO-{24000 + n}", "client": rnd.choice(CLIENTS),
                  "store_name": store, "location": loc,
                  "address": f"{rnd.randint(100, 9999)} Main St, {loc}",
                  "job_type": rnd.choice(JOBS), "handler": rnd.choice(PEOPLE),
                  "tech": rnd.choice(TECHS), "assignees": rnd.choice(TECHS),
                  "wo_status": status, "date_in": date_in.isoformat(),
                  "eta": (date_in + timedelta(days=rnd.randint(1, 7))).isoformat(),
                  "direct_cost": cost, "target_markup": markup, "approved_nte": nte,
                  "invoice_status": "Not invoiced", "payment_terms": rnd.choice([15, 30, 45])}
            if status == "Awaiting parts":
                wo["parts_status"] = "Compressor ordered"
            if status == "Completed":
                sent = date_in + timedelta(days=rnd.randint(3, 14))
                revenue = round(cost * (1 + markup) * rnd.uniform(0.9, 1.1), 2)
                paid = rnd.choice([0, revenue, revenue, round(revenue / 2, 2)])
                wo.update(invoice_number=f"INV-{5000 + n}", invoice_sent_date=sent.isoformat(),
                          invoice_revenue=revenue, amount_received=paid,
                          overhead=round(revenue * 0.08, 2), admin_fees=round(revenue * 0.03, 2),
                          other_expenses=rnd.choice([0, 0, 25.0, None]),
                          payment_method=rnd.choice(["ACH", "Check", "Credit card"]),
                          invoice_status="Paid" if paid >= revenue else
                          ("Partially paid" if paid else "Sent"))
            db.insert(wo, action="Created (demo data)")
        print("Seeded 40 demo work orders.")


if __name__ == "__main__":
    main()
