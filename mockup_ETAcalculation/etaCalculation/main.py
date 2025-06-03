from datetime import datetime, timedelta

#lookup table: (departure, destination): (eta_duration, delivery_duration)
etd_eta_table = {
    ("NINGBO", "TAULOV"): (54, 61),
    ("NINGBO", "FOS SUR MER"): (20, 25),
    # ETD ETA table from NS
}

# constant for the missing routes
MISSING_ROUTE_DURATION = 500

# an example of a current PO data
record = {
    "[SCH] ETD": datetime.strptime("20/03/2025", "%d/%m/%Y"),
    "[SCH] ETA port": datetime.strptime("30/04/2025", "%d/%m/%Y"),
    "[SCH] Delivery date": datetime.strptime("05/05/2025", "%d/%m/%Y"),
    "[SCH] ETA port set": False,
    "[SCH] delivery confirmed": False,
    "[SCH] Departure port": "NINGBO",
    "[SCH] Destination port": "TAULOV"
}

def lookup_durations(departure, destination):
    return etd_eta_table.get((departure, destination), (MISSING_ROUTE_DURATION, MISSING_ROUTE_DURATION))

def calculate_eta(current_eta, current_etd, eta_port_set, eta_duration_days):
    return current_eta if eta_port_set else current_etd + timedelta(days=eta_duration_days)

def calculate_delivery_date(delivery_confirmed, current_delivery_date, etd, delivery_duration_days):
    return current_delivery_date if delivery_confirmed else etd + timedelta(days=delivery_duration_days)


def main():
    # Get user input for dates — allow empty input
    new_etd_input_str = input("ETD date received (DD/MM/YYYY) [leave empty to skip]: ").strip()
    new_eta_input_str = input("ETA date received (DD/MM/YYYY) [leave empty to skip]: ").strip()

    # Default to current values
    current_etd = record["[SCH] ETD"]
    current_eta = record["[SCH] ETA port"]
    eta_port_set = record["[SCH] ETA port set"]

    # Parse ETD input if provided
    if new_etd_input_str:
        try:
            current_etd = datetime.strptime(new_etd_input_str, "%d/%m/%Y")
        except ValueError:
            print("Invalid ETD date format. Please use DD/MM/YYYY.")
            return

    # Parse ETA input if provided
    if new_eta_input_str:
        try:
            new_eta_input = datetime.strptime(new_eta_input_str, "%d/%m/%Y")
            if new_eta_input == current_eta:
                eta_port_set = True
            else:
                eta_port_set = True
                current_eta = new_eta_input
        except ValueError:
            print("Invalid ETA date format. Please use DD/MM/YYYY.")
            return

    current_delivery_date = record["[SCH] Delivery date"]
    delivery_confirmed = record["[SCH] delivery confirmed"]
    departure_port = record["[SCH] Departure port"]
    destination_port = record["[SCH] Destination port"]

    eta_duration, delivery_duration = lookup_durations(departure_port, destination_port)

    new_etd = current_etd
    new_eta = calculate_eta(current_eta,current_etd, eta_port_set, eta_duration)
    new_delivery_date = calculate_delivery_date(delivery_confirmed, current_delivery_date, new_etd, delivery_duration)

    # Update the record
    record["[SCH] ETD"] = new_etd
    record["[SCH] ETA port"] = new_eta
    record["[SCH] Delivery date"] = new_delivery_date
    record["[SCH] ETA port set"] = eta_port_set

    print("\n Updated record:")
    for key in [
        "[SCH] Departure port",
        "[SCH] Destination port",
        "[SCH] ETD",
        "[SCH] ETA port",
        "[SCH] ETA port set",
        "[SCH] Delivery date",
        "[SCH] delivery confirmed"
    ]:
        value = record[key]
        if isinstance(value, datetime):
            print(f"{key}: {value.strftime('%d/%m/%Y')}")
        else:
            print(f"{key}: {value}")



if __name__ == "__main__":
    main()
