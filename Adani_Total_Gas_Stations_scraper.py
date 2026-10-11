import requests
import pandas as pd
import time

# ---------------- Configuration ---------------- #

HEADERS = {
    "User-Agent": "Mozilla/5.0",
    "Referer": "https://www.adanigas.com/en/cng/cng-locate",
    "Origin": "https://www.adanigas.com",
    "X-Requested-With": "XMLHttpRequest"
}

AREA_API = "https://www.adanigas.com/api/AdaniGas/CngArea"
STATION_API = "https://www.adanigas.com/api/AdaniGas/CNGStations"

cities = [
    "Ahmedabad","Ahwa Dang","Akola","Alirajpur","Amravati","Annupur",
    "Anuppur","Balaikheda","Balasore","Barpeta","Barwala","Barwani",
    "Bhadrak","Bhandara","Bhilwara","Bhind","Bhind , M.P","Bhiwani",
    "Bilaspur","Botad","Bundi","Burhanpur","charkhi Dadri",
    "Charkhi Dadri","Chhatarpur","CHIRANG","Chittorgarh","Cuddalore",
    "Datia","Dhariyawad","Durg","Faridabad","FARIDABAD","Garhwa",
    "Gondia","Gumla","GUMLA","Halvad","Harda","hingoli","Hingoli",
    "Jalandhar","Jalaun","Jhansi","Khandwa","Khargone","KHARGONE",
    "Kheda","Khunti","Khurja","KHURJA","Kokrajhar","Koraput",
    "Latehar","Lohardaga","Lunawada","Mahasamund","Mahendragarh",
    "Mahendragrh","Mahisagar","Mahuva, Surat","Mayiladuthurai",
    "Mayurbhanj","MEWAT","Morbi","Nabarangapur","Nagaon",
    "Nandurbar","Narnaul","Navsari","Nuh","Palwal","PALWAL",
    "Panna","Porbandar","PORBANDAR","Pratapgarh","Rajnandgaon",
    "Simdega","Srirampur","Surat","Surendranagar","Tamil Nadu",
    "Thiruvarur","Tikamgarh","Tiruppur","Udaipur","UDAIPUR",
    "Udupi","UDUPI","Vadodara","Vansda","Wani","Yavatmal"
]

# Remove duplicate city names while preserving order
cities = list(dict.fromkeys(cities))

# ---------------- Scraping ---------------- #

records = []

for city in cities:

    print(f"\nProcessing: {city}")

    try:
        area_response = requests.get(
            AREA_API,
            params={"city": city},
            headers=HEADERS,
            timeout=30
        )

        if area_response.status_code != 200:
            print("Failed:", city)
            continue

        areas = area_response.json()

        print(f"Areas: {len(areas)}")

        for area in areas:

            area_name = area["Value"]

            payload = {
                "city": city,
                "area": area_name
            }

            station_response = requests.post(
                STATION_API,
                json=payload,
                headers=HEADERS,
                timeout=30
            )

            if station_response.status_code != 200:
                continue

            stations = station_response.json()

            print(f"   {area_name:35} -> {len(stations)}")

            records.extend(stations)

            time.sleep(0.15)

    except Exception as e:
        print(city, e)

# ---------------- Export ---------------- #

df = pd.DataFrame(records)

print("\nRaw Records:", len(df))

df.drop_duplicates(subset="LocNo", inplace=True)

print("Unique Stations:", len(df))

df["Station Name"] = df["AddrLine1"]

df["State"] = df["AddrLine5"]

df["Address"] = (
    df["AddrLine1"].fillna("") + ", " +
    df["AddrLine2"].fillna("") + ", " +
    df["AddrLine3"].fillna("") + ", " +
    df["AddrLine4"].fillna("") + ", " +
    df["AddrLine5"].fillna("")
)

df["Google Maps"] = (
    "https://www.google.com/maps?q="
    + df["Latitude"].astype(str)
    + ","
    + df["Longitude"].astype(str)
)

columns = [
    "LocNo",
    "Station Name",
    "Address",
    "NearArea",
    "City",
    "State",
    "PinCode",
    "MobileNo",
    "Latitude",
    "Longitude",
    "Google Maps"
]

df = df[columns]

df.sort_values(["State", "City", "Station Name"], inplace=True)

df.to_excel("Adani_Total_Gas_Stations.xlsx", index=False)

print("\nDone!")
print("Excel saved as: Adani_Total_Gas_Stations.xlsx")
print("Total Stations:", len(df))