# TS Booking Checking

Standalone TS Booking Control web app for direct-voyage monitoring, transfer monitoring, and transfer-port/date-range searches.

The published page loads its packaged Booking and schedule snapshot first, then attempts to refresh Daily Booking data from `BOBWZW2/data-base`.

## Local preview

Serve this repository with any static HTTP server and open `index.html`.

## Data rebuild

Run `scripts/generate_booking_data.py` to rebuild `data/booking-data.js` and `data/booking-data.json` from the current Booking and schedule workbooks.
