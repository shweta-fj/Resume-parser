"""
Generate fake resumes (tech + sales) for load-testing Meilisearch.
Outputs a JSON array file ready for bulk POST to /indexes/<uid>/documents.

Usage:
    python3 generate_resumes.py --count 500000 --out resumes.json
"""

import argparse
import json
import random
from faker import Faker

fake = Faker()

TECH_ROLES = [
    "Software Engineer", "Backend Developer", "Frontend Developer",
    "Full Stack Developer", "DevOps Engineer", "Data Scientist",
    "Data Analyst", "ML Engineer", "QA Engineer", "Mobile Developer",
    "Cloud Architect", "Site Reliability Engineer", "Security Engineer",
    "Database Administrator", "Systems Engineer",
]

SALES_ROLES = [
    "Sales Executive", "Account Manager", "Business Development Manager",
    "Sales Development Representative", "Regional Sales Manager",
    "Key Account Manager", "Inside Sales Representative",
    "Enterprise Sales Manager", "Sales Operations Analyst",
    "Channel Sales Manager", "Customer Success Manager",
    "Territory Sales Manager", "Sales Consultant",
]

TECH_SKILLS = [
    "Python", "JavaScript", "TypeScript", "React", "Node.js", "Django",
    "Flask", "PostgreSQL", "MySQL", "MongoDB", "Docker", "Kubernetes",
    "AWS", "GCP", "Azure", "Redis", "GraphQL", "REST APIs", "CI/CD",
    "Git", "Java", "Spring Boot", "Go", "Rust", "C++", "Terraform",
    "Machine Learning", "TensorFlow", "PyTorch", "SQL", "Linux",
    "Microservices", "Kafka", "Elasticsearch", "Meilisearch",
]

SALES_SKILLS = [
    "B2B Sales", "CRM (Salesforce)", "Lead Generation", "Cold Calling",
    "Negotiation", "Account Management", "Pipeline Management",
    "Upselling", "Client Relationship Management", "HubSpot",
    "Sales Forecasting", "Market Research", "Contract Negotiation",
    "Customer Retention", "Territory Management", "SaaS Sales",
    "Consultative Selling", "Zoho CRM", "Outbound Prospecting",
    "Presentation Skills",
]

TECH_COMPANIES = [
    "Infosys", "TCS", "Wipro", "HCL Technologies", "Tech Mahindra",
    "Zoho", "Freshworks", "Razorpay", "Swiggy", "Zomato", "Flipkart",
    "Paytm", "PhonePe", "CRED", "Postman", "BrowserStack", "Amazon",
    "Microsoft", "Google", "Adobe", "Uber", "OYO", "Meesho",
]

SALES_COMPANIES = [
    "Reliance Retail", "Byju's", "Vedantu", "Policybazaar", "Naukri.com",
    "IndiaMART", "Urban Company", "Lenskart", "Nykaa", "boAt",
    "Cars24", "Pepperfry", "Zoho", "Freshworks", "Amazon", "HDFC Bank",
    "ICICI Bank", "Asian Paints", "Hindustan Unilever", "Godrej",
]

INDIAN_CITIES = [
    "Bengaluru", "Hyderabad", "Pune", "Gurugram", "Noida", "Delhi",
    "Mumbai", "Chennai", "Kolkata", "Ahmedabad", "Jaipur", "Kochi",
    "Chandigarh", "Indore", "Lucknow", "Greater Noida", "Nagpur",
]

TECH_DEGREES = [
    "B.Tech in Computer Science", "B.E. in Information Technology",
    "M.Tech in Data Science", "B.Sc in Computer Science",
    "MCA", "B.Tech in Electronics", "M.Sc in Data Science",
    "B.E. in Computer Engineering",
]

SALES_DEGREES = [
    "MBA in Marketing", "MBA in Sales & Operations", "B.Com", "BBA",
    "B.A. Economics", "MBA in Business Administration",
    "B.Com in Marketing", "PGDM in Sales",
]


def gen_tech_resume(idx):
    n_skills = random.randint(4, 8)
    return {
        "id": idx,
        "category": "tech",
        "name": fake.name(),
        "email": fake.unique.email(),
        "phone": fake.phone_number(),
        "city": random.choice(INDIAN_CITIES),
        "role": random.choice(TECH_ROLES),
        "years_experience": random.randint(0, 18),
        "current_company": random.choice(TECH_COMPANIES),
        "skills": random.sample(TECH_SKILLS, n_skills),
        "education": random.choice(TECH_DEGREES),
        "expected_salary_lpa": round(random.uniform(3, 60), 1),
        "summary": fake.paragraph(nb_sentences=3),
        "notice_period_days": random.choice([0, 15, 30, 60, 90]),
    }


def gen_sales_resume(idx):
    n_skills = random.randint(4, 8)
    return {
        "id": idx,
        "category": "sales",
        "name": fake.name(),
        "email": fake.unique.email(),
        "phone": fake.phone_number(),
        "city": random.choice(INDIAN_CITIES),
        "role": random.choice(SALES_ROLES),
        "years_experience": random.randint(0, 20),
        "current_company": random.choice(SALES_COMPANIES),
        "skills": random.sample(SALES_SKILLS, n_skills),
        "education": random.choice(SALES_DEGREES),
        "expected_salary_lpa": round(random.uniform(2.5, 40), 1),
        "summary": fake.paragraph(nb_sentences=3),
        "notice_period_days": random.choice([0, 15, 30, 60, 90]),
    }


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--count", type=int, default=500_000)
    parser.add_argument("--out", type=str, default="resumes.json")
    parser.add_argument("--split", type=float, default=0.5,
                         help="fraction that should be tech (rest is sales)")
    args = parser.parse_args()

    tech_count = int(args.count * args.split)
    sales_count = args.count - tech_count

    print(f"Generating {tech_count} tech + {sales_count} sales = {args.count} resumes...")

    with open(args.out, "w") as f:
        f.write("[")
        first = True
        idx = 1

        for _ in range(tech_count):
            if not first:
                f.write(",")
            json.dump(gen_tech_resume(idx), f)
            first = False
            idx += 1
            if idx % 50000 == 0:
                print(f"  ...{idx} generated")

        for _ in range(sales_count):
            if not first:
                f.write(",")
            json.dump(gen_sales_resume(idx), f)
            first = False
            idx += 1
            if idx % 50000 == 0:
                print(f"  ...{idx} generated")

        f.write("]")

    print(f"Done. Wrote {args.count} resumes to {args.out}")


if __name__ == "__main__":
    main()
