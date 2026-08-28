"""Test ZIP that includes Airtel 2401: IPv6 records."""
import os, zipfile, tempfile

SUBSCRIBER = """<html><body>
<h1>GOOGLE SUBSCRIBER INFORMATION</h1>
<p>Google Account ID: 9988776655</p>
<p>Name: Airtel Tester</p>
<p>e-Mail: airtel@example.com</p>
<p>Created on: 2024-01-01 00:00:00 UTC</p>
<table border="1">
<thead>
<tr><th>Timestamp</th><th>IP Address</th><th>Activity Type</th></tr>
</thead>
<tbody>
<tr><td>2025-12-15 10:00:00 UTC</td><td>2401:4900:abc:1::1</td><td>Successful login</td></tr>
<tr><td>2025-12-15 11:30:00 UTC</td><td>2401:4900:abc:1::2</td><td>Failed login</td></tr>
<tr><td>2025-12-15 12:45:00 UTC</td><td>203.0.113.42</td><td>Failed login</td></tr>
<tr><td>2025-12-15 13:50:00 UTC</td><td>198.51.100.7</td><td>Successful login</td></tr>
</tbody>
</table>
</body></html>
"""

with zipfile.ZipFile(
    os.path.join(tempfile.gettempdir(), "airtel_export.zip"),
    "w",
    zipfile.ZIP_DEFLATED,
) as z:
    z.writestr("subscriber.html", SUBSCRIBER)
print("Created airtel_export.zip")