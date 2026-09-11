import requests

data = {
    "year": "2026",
    "month": "10",
    "prior_fiscal_days": "183"
}

# Fill commodity targets
for i in range(13):
    data[f"commodity_target_{i}"] = "100"
    data[f"current_commodity_{i}"] = "1000"
    data[f"prior_commodity_{i}"] = "900"
    data[f"prior_full_commodity_{i}"] = "1100"

# Fill division targets
for i in range(7):
    data[f"division_target_{i}"] = "50"
    data[f"current_division_{i}"] = "500"
    data[f"prior_division_{i}"] = "450"
    data[f"prior_full_division_{i}"] = "550"

# Fill prior daily grids (13 rows, 31 cols)
for r in range(13):
    for d in range(1, 32):
        data[f"prior_daily_commodity_{r}_{d}"] = "10"

# Fill prior daily division grids (7 rows, 31 cols)
for r in range(7):
    for d in range(1, 32):
        data[f"prior_daily_division_{r}_{d}"] = "5"

response = requests.post("http://127.0.0.1:5050/generate", data=data)

if response.status_code == 200:
    print("Success. Response contains:", response.text[:200])
else:
    print("Failed with status:", response.status_code)
    print(response.text)
