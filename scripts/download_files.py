"""
download_files.py
Downloads all 21 AllView Excel files from personal OneDrive.
Uses requests with session handling for personal OneDrive "Anyone" links.
"""

import os
import sys
import time

try:
    import requests
except ImportError:
    print("Installing requests...")
    os.system("pip install requests -q")
    import requests

OUTPUT_DIR = './data'
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ── OneDrive share links ──────────────────────────────────────────────────────
FILE_LINKS = {
    '5-star_Tracker.xlsx':
        'https://1drv.ms/x/c/7e1a6235327bfe3a/IQAvLLfT9zjRTImLtqAqU4XmAdXCr5WOMjrXNS0XfHOrVdA?e=VC0piN',
    'BDM_Data.xlsx':
        'https://1drv.ms/x/c/7e1a6235327bfe3a/IQDVQEcz0a17RJsjnPt7_yeTARm1tBN2tK7CikL1NZ0NJzI?e=mt3OYJ',
    'Email_Analytics.xlsx':
        'https://1drv.ms/x/c/7e1a6235327bfe3a/IQCHXps3cxMmRIKJzLmBq_m7AS61tqdwQbgETz-GTqybFEM?e=4Tu5pN',
    'guest_card_inquiries-Performance.xlsx':
        'https://1drv.ms/x/c/7e1a6235327bfe3a/IQCs1GTUWBlsSqrP4HXp13z3AZO69QCRjzcR_JncmTOBrXA?e=AYMpfZ',
    'LR_Performance.xlsx':
        'https://1drv.ms/x/c/7e1a6235327bfe3a/IQB9pg6SFjxzTIRogWUauv-PAb0So0smpLhPWWfw5Rho1j0?e=ZeDjAr',
    'New_NPS_Data.xlsx':
        'https://1drv.ms/x/c/7e1a6235327bfe3a/IQCUhRLm953gQI0G-YTe3hgaAdr0q1oRgbfUBIwmEKnlbd8?e=uNeOhS',
    'New_Portafolios_Structure_-_Effective_Feb_17.xlsx':
        'https://1drv.ms/x/c/7e1a6235327bfe3a/IQAo3N_EJHgGT4-62WJcnVupASwUfG5Cw7ThoNaefXBIZIg?e=Xa1IR9',
    'NPS_Data.xlsx':
        'https://1drv.ms/x/c/7e1a6235327bfe3a/IQB53gToggnKRqdIf1Dr4_1ZAaWSp_fCLJnTrDAfTbn0F5Q?e=eX5LYh',
    'Owner_Concessions.xlsx':
        'https://1drv.ms/x/c/7e1a6235327bfe3a/IQBqWftit5J9S4JBTWGfIJvMAWwL_R_v_vLpMAPKnqUvQko?e=dLt6Ir',
    'Phone_Reporting.xlsx':
        'https://1drv.ms/x/c/7e1a6235327bfe3a/IQAJ6Jjo64gAQ7bix0oKMYW0ATWJ2Uc71L8jAeqK3Viq2A0?e=vKrDeX',
    'property_directory-Term.xlsx':
        'https://1drv.ms/x/c/7e1a6235327bfe3a/IQBlBPyMpSKKTJdHnjrR2hrwAS5mRSG9HrJcUkoHdafji6U?e=pqg4u5',
    'Rent_Roll.xlsx':
        'https://1drv.ms/x/c/7e1a6235327bfe3a/IQAGNQRdDk5gTLjX0vsp9eTpAXdhK3ekMEzFKDMORqNyWoc?e=vPUpHJ',
    'Rental_Vacancy.xlsx':
        'https://1drv.ms/x/c/7e1a6235327bfe3a/IQCihKdjblPET4h7TSthEqGqAdXrQNgffqB9PYYRdvvpVl4?e=NSecck',
    'rental_applications-Performance.xlsx':
        'https://1drv.ms/x/c/7e1a6235327bfe3a/IQC8OlURHPRyQbYNPNEMFilMAW543mzV_Tluy9szSBWytfU?e=Gbva4w',
    'tenant_tickler_-_Move-in-out.xlsx':
        'https://1drv.ms/x/c/7e1a6235327bfe3a/IQDdUQ5FWCdTT5fMAthLKT1-AVfT_btZ-Ro6L1LkcQrRXdI?e=VT1Ryg',
    'tenant_unpaid_charges_summary.xlsx':
        'https://1drv.ms/x/c/7e1a6235327bfe3a/IQB4rfITXlxzSpHkLVCF9at0AZsm6xUgWlVnveNu181MkHU?e=dIgxbH',
    'Unit_Turn_Details.xlsx':
        'https://1drv.ms/x/c/7e1a6235327bfe3a/IQBG8HrtVhdBSYuDJ7e_nG6LAQfwiGUOFMVvlIHI-rUaa7s?e=BZ0LNK',
    'WorkOrders_Report.xlsx':
        'https://1drv.ms/x/c/7e1a6235327bfe3a/IQDKMJxWJS8gTpZhbAc_MqBFAeHL9Q87yrA7wBq7uAx1QI8?e=7h9x6v',
    'Yelp-Google_Reviews.xlsx':
        'https://1drv.ms/x/c/7e1a6235327bfe3a/IQDQXfoH5ci9Qr8DnAloTOlHAf87jQg7da2vziFVdUDEua8?e=0DDWE0',
    'Active_Listings_-_Performance.xlsx':
        'https://1drv.ms/x/c/7e1a6235327bfe3a/IQDF39GDg437SbKrf3uS2ZtEAbxNAO3I1Ga6qoJ26xuVfvU?e=j807ko',
    'Off_Market_Listings_-_Performance__Historical_Data_.xlsx':
        'https://1drv.ms/x/c/7e1a6235327bfe3a/IQBk7NK4-7NCTJ5Hg9zIOMpXAbfm015W5Xtwh01I0u7ptFA?e=2GUtbe',
}


def get_download_url(share_url, session):
    """
    Convert a personal OneDrive share URL to a direct download URL.
    Follows the full redirect chain using a session with cookies.
    """
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
                      'AppleWebKit/537.36 (KHTML, like Gecko) '
                      'Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'en-US,en;q=0.9',
    }

    # Step 1: Follow the short URL to get the full OneDrive URL
    resp = session.get(share_url, headers=headers, allow_redirects=True, timeout=30)
    final_url = resp.url

    # Step 2: Try adding download=1 to trigger file download
    if '?' in final_url:
        download_url = final_url + '&download=1'
    else:
        download_url = final_url + '?download=1'

    return download_url


def download_file(filename, share_url, session, output_dir, retries=3):
    """Download a single file using a persistent session."""
    out_path = os.path.join(output_dir, filename)

    for attempt in range(1, retries + 1):
        try:
            download_url = get_download_url(share_url, session)

            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
                              'AppleWebKit/537.36 (KHTML, like Gecko) '
                              'Chrome/120.0.0.0 Safari/537.36',
            }

            resp = session.get(
                download_url,
                headers=headers,
                allow_redirects=True,
                timeout=120,
                stream=True
            )

            if resp.status_code != 200:
                raise ValueError(f'HTTP {resp.status_code}')

            content = resp.content

            # Validate Excel file (PK zip header)
            if len(content) < 4 or content[:2] != b'PK':
                # Try the content-type header
                ct = resp.headers.get('Content-Type', '')
                raise ValueError(
                    f'Not a valid Excel file. '
                    f'Content-Type: {ct}, '
                    f'Size: {len(content)} bytes, '
                    f'Header: {content[:8]}'
                )

            with open(out_path, 'wb') as f:
                f.write(content)

            print(f'  OK  {filename:<55} {len(content)//1024:>6} KB')
            return True

        except Exception as e:
            if attempt < retries:
                time.sleep(5)
                print(f'  ..  {filename} — retry {attempt}/{retries}: {e}')
            else:
                print(f'  XX  {filename} — FAILED: {e}')
                return False


def main():
    print('=' * 65)
    print('AllView Dashboard — OneDrive File Downloader')
    print('Personal OneDrive with session-based download')
    print('=' * 65)
    print(f'\nDownloading {len(FILE_LINKS)} files to {OUTPUT_DIR}/\n')

    # Use a persistent session so cookies carry across requests
    session = requests.Session()

    # Prime the session by visiting OneDrive first
    try:
        session.get(
            'https://onedrive.live.com',
            headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0.0.0'},
            timeout=10
        )
    except Exception:
        pass  # OK if this fails, proceed anyway

    succeeded, failed = [], []

    for filename, share_url in FILE_LINKS.items():
        ok = download_file(filename, share_url, session, OUTPUT_DIR)
        (succeeded if ok else failed).append(filename)

    print(f'\n{"=" * 65}')
    print(f'Result: {len(succeeded)}/{len(FILE_LINKS)} files downloaded')

    if failed:
        print(f'\nFailed ({len(failed)}):')
        for f in failed:
            print(f'  - {f}')
        sys.exit(1)
    else:
        print('All 21 files ready. Running build.py next...')
    print('=' * 65)


if __name__ == '__main__':
    main()
