"""
download_files.py
Downloads all 19 AllView Excel files from OneDrive using individual share links.
No authentication required — links are publicly shared.
Saves files to ./data/ folder for build.py to process.
"""

import os
import sys
import base64
import urllib.request
import urllib.error

OUTPUT_DIR = './data'
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ── OneDrive share links for each file ───────────────────────────────────────
# To update a link: replace the URL for that file below.
# Links are generated from OneDrive → right-click file → Share → Copy link.
FILE_LINKS = {
    '5-star_Tracker.xlsx':
        'https://1drv.ms/x/c/7e1a6235327bfe3a/IQAvLLfT9zjRTImLtqAqU4XmAdXCr5WOMjrXNS0XfHOrVdA?e=fKwDI3',

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
}


def share_link_to_download_url(share_url):
    """Convert a OneDrive share link to a direct download URL."""
    b64 = base64.b64encode(share_url.encode('utf-8')).decode('utf-8')
    b64 = b64.rstrip('=').replace('/', '_').replace('+', '-')
    token = 'u!' + b64
    return f'https://api.onedrive.com/v1.0/shares/{token}/root/content'


def download_file(filename, share_url, output_dir, retries=3):
    """Download a single file from a OneDrive share link."""
    download_url = share_link_to_download_url(share_url)
    out_path     = os.path.join(output_dir, filename)

    for attempt in range(1, retries + 1):
        try:
            req = urllib.request.Request(
                download_url,
                headers={'User-Agent': 'Mozilla/5.0 (AllView Dashboard Builder)'}
            )
            with urllib.request.urlopen(req, timeout=60) as resp:
                content = resp.read()

            # Validate it's an Excel file (PK zip header)
            if len(content) < 4 or content[:2] != b'PK':
                raise ValueError(
                    f'Not a valid Excel file ({len(content)} bytes, '
                    f'header: {content[:4]})'
                )

            with open(out_path, 'wb') as f:
                f.write(content)

            print(f'  OK  {filename:<55} {len(content)//1024:>6} KB')
            return True

        except urllib.error.HTTPError as e:
            if attempt < retries:
                import time; time.sleep(3)
                print(f'  ..  {filename} — HTTP {e.code}, retry {attempt}/{retries}')
            else:
                print(f'  XX  {filename} — FAILED: HTTP {e.code}')
                print(f'      Tip: re-share the file in OneDrive and update its URL in this script.')
                return False

        except Exception as e:
            if attempt < retries:
                import time; time.sleep(3)
                print(f'  ..  {filename} — {e}, retry {attempt}/{retries}')
            else:
                print(f'  XX  {filename} — FAILED: {e}')
                return False


def main():
    print('=' * 65)
    print('AllView Dashboard — OneDrive File Downloader')
    print('Method: public share links (no authentication needed)')
    print('=' * 65)
    print(f'\nDownloading {len(FILE_LINKS)} files to {OUTPUT_DIR}/\n')

    succeeded, failed = [], []

    for filename, share_url in FILE_LINKS.items():
        ok = download_file(filename, share_url, OUTPUT_DIR)
        (succeeded if ok else failed).append(filename)

    print(f'\n{"=" * 65}')
    print(f'Result: {len(succeeded)}/{len(FILE_LINKS)} files downloaded')

    if failed:
        print(f'\nFailed ({len(failed)}):')
        for f in failed:
            print(f'  - {f}')
        print('\nFix: in OneDrive, right-click each failed file → Share → Copy link')
        print('     then update FILE_LINKS in this script with the new URL.')
        sys.exit(1)
    else:
        print('All 19 files ready. Run build.py next.')
    print('=' * 65)


if __name__ == '__main__':
    main()
