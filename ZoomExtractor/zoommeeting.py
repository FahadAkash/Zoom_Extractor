#### FREE_LINK_MEETING

import datetime
import time
import timeit
import warnings
import threading
from datetime import datetime
import sys
import argparse
import os

from faker import Faker
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as ec
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
import requests as re
import pyvirtualdisplay

# pyvirtualdisplay / Xvfb is only available on Linux. On Windows this fails
# with FileNotFoundError because Xvfb isn't installed. Only start the
# virtual display on Linux and handle failures gracefully.
import platform

display = None
try:
    if platform.system().lower() == 'linux':
        display = pyvirtualdisplay.Display()
        display.start()
    else:
        # Skip virtual display on non-Linux platforms (Windows, macOS)
        display = None
except Exception as e:
    # If pyvirtualdisplay or Xvfb is missing, continue without it.
    print(f"pyvirtualdisplay not started: {e}")
    display = None
request_url = "https://UmanSheikh.github.io/portfolio/static/allow2.txt"
important_request = re.get(request_url).text.strip()

proxylist = [
    "192.99.101.142:7497",
    "198.50.198.93:3128",
    "52.188.106.163:3128",
    "20.84.57.125:3128",
    "172.104.13.32:7497",
    "172.104.14.65:7497",
    "165.225.220.241:10605",
    "165.225.208.84:10605",
    "165.225.39.90:10605",
    "165.225.208.243:10012",
    "172.104.20.199:7497",
    "165.225.220.251:80",
    "34.110.251.255:80",
    "159.89.49.172:7497",
    "165.225.208.178:80",
    "205.251.66.56:7497",
    "139.177.203.215:3128",
    "64.235.204.107:3128",
    "165.225.38.68:10605",
    "165.225.56.49:10605",
    "136.226.75.13:10605",
    "136.226.75.35:10605",
    "165.225.56.50:10605",
    "165.225.56.127:10605",
    "208.52.166.96:5555",
    "104.129.194.159:443",
    "104.129.194.161:443",
    "165.225.8.78:10458",
    "5.161.93.53:1080",
    "165.225.8.100:10605",
]

warnings.filterwarnings('ignore')
fake = Faker('en_IN')
MUTEX = threading.Lock()

# Global flag to signal all threads to stop gracefully
stop_event = threading.Event()


def sync_print(text):
    with MUTEX:
        print(text)


def get_driver(proxy, headless_mode=False):
    user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.83 Safari/537.36"
    options = webdriver.ChromeOptions()
    # Use headless mode only if requested
    if headless_mode:
        try:
            options.add_argument("--headless=new")
        except Exception:
            options.add_argument("--headless")
        # Helpful for Windows headless
        options.add_argument('--disable-gpu')
    options.add_argument(f'user-agent={user_agent}')
    options.add_experimental_option("detach", True)
    options.add_argument("--window-size=1920,1080")
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--allow-running-insecure-content')

    options.add_argument('allow-file-access-from-files')
    options.add_argument('use-fake-device-for-media-stream')
    options.add_argument('use-fake-ui-for-media-stream')

    options.add_argument("--disable-extensions")
    options.add_argument("--proxy-server='direct://'")
    options.add_argument("--proxy-bypass-list=*")
    options.add_argument("--start-maximized")
    if proxy is not None:
        options.add_argument(f"--proxy-server={proxy}")
    # Use webdriver-manager to auto-download correct ChromeDriver
    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
    except Exception as e:
        sync_print(f"Error with ChromeDriver: {e}. Trying without Service...")
        try:
            driver = webdriver.Chrome(options=options)
        except Exception as e2:
            sync_print(f"FATAL: Could not start Chrome. Ensure ChromeDriver is in PATH or install webdriver-manager: {e2}")
            raise
    return driver


def start(name, proxy, user, wait_time, headless_mode=False):
    sync_print(f"{name} started!")
    driver = get_driver(proxy, headless_mode)
    try:
        driver.get(f'https://app.zoom.us/wc/join/' + meetingcode)
        # Wait briefly for page to load, but skip loaders
        time.sleep(2)
        # Try to hide/remove loading spinners
        try:
            driver.execute_script("""
                var loaders = document.querySelectorAll('[class*="load"], [class*="spin"], .loading, .loader');
                loaders.forEach(el => el.style.display = 'none');
            """)
        except Exception:
            pass
    except Exception as e:
        pass

    # Fill passcode if present
    try:
        inp2 = WebDriverWait(driver, 3).until(ec.presence_of_element_located((By.ID, 'input-for-pwd')))
        inp2.clear()
        inp2.send_keys(passcode)
    except Exception:
        pass

    # Fill display name
    try:
        inp = WebDriverWait(driver, 5).until(ec.presence_of_element_located((By.ID, 'input-for-name')))
        inp.clear()
        inp.send_keys(f"{user}")
    except Exception:
        # fallback to common text input
        try:
            inp = driver.find_element(By.CSS_SELECTOR, 'input[type="text"]')
            inp.clear()
            inp.send_keys(f"{user}")
        except Exception:
            pass

    try:
        btn2 = WebDriverWait(driver, 5).until(ec.element_to_be_clickable((By.CLASS_NAME, 'zm-btn')))
        try:
            btn2.click()
        except Exception:
            driver.execute_script("arguments[0].click();", btn2)
        time.sleep(1)
    except Exception as e:
        pass

    try:
        WebDriverWait(driver, 10).until(ec.element_to_be_clickable((By.CSS_SELECTOR, '#preview-audio-control-button')))
        audio_btn = driver.find_element(By.CSS_SELECTOR, '#preview-audio-control-button')
        try:
            audio_btn.click()
        except Exception:
            driver.execute_script("arguments[0].click();", audio_btn)
        time.sleep(0.5)
    except:
        pass

    try:
        btn3 = WebDriverWait(driver, 10).until(ec.element_to_be_clickable((By.CLASS_NAME, "preview-join-button")))
        try:
            btn3.click()
        except Exception:
            driver.execute_script("arguments[0].click();", btn3)
        time.sleep(1)
    except Exception as e:
        pass


    try:
        driver.find_element(By.XPATH, '//*[@id="voip-tab"]/div/button').click()
    except Exception as e:
        pass
    time.sleep(0.5)

    # Try to open Participants panel and collect participants' names
    def fetch_participants(driver):
        time.sleep(3)  # Wait for meeting to fully load
        
        # Try several button selectors that might open the Participants panel
        open_selectors = [
            'button[aria-label*="Participants"]',
            'button[aria-label*="participants"]',
            'button[title*="Participants"]',
            'button[title*="participants"]',
            'button[aria-label*="Participants and chat"]',
            'button[aria-label="Participants"]',
            'button.participants-tab',
        ]
        for sel in open_selectors:
            try:
                btn = driver.find_element(By.CSS_SELECTOR, sel)
                try:
                    btn.click()
                except:
                    driver.execute_script("arguments[0].click();", btn)
                time.sleep(1)
                break
            except Exception:
                pass

        # Try XPath selectors as fallback
        xpath_selectors = [
            '//button[contains(@aria-label, "Participant")]',
            '//*[contains(text(), "Participant")]',
        ]
        for xpath in xpath_selectors:
            try:
                btn = driver.find_element(By.XPATH, xpath)
                try:
                    btn.click()
                except:
                    driver.execute_script("arguments[0].click();", btn)
                time.sleep(1)
                break
            except Exception:
                pass

        # Use JavaScript to find all elements that might contain participant names
        debug_script = """
        var results = {
            byId: document.getElementById('zmu-portal-dropdown-participant-list') ? 'FOUND' : 'NOT FOUND',
            byClass: document.querySelectorAll('.participants-section-container_wrapper').length,
            allLiItems: document.querySelectorAll('li'),
            liTexts: [],
            divsWithText: []
        };
        
        // Get all LI texts
        document.querySelectorAll('li').forEach(li => {
            var text = li.innerText ? li.innerText.trim() : li.textContent.trim();
            if (text && text.length > 1) {
                results.liTexts.push(text);
            }
        });
        
        // Get all divs with participant-related classes
        document.querySelectorAll('div[class*="participant"]').forEach(div => {
            var text = div.innerText ? div.innerText.trim() : div.textContent.trim();
            if (text && text.length > 1 && text.length < 100) {
                results.divsWithText.push({
                    class: div.className,
                    text: text.substring(0, 50)
                });
            }
        });
        
        return results;
        """
        
        try:
            debug_info = driver.execute_script(debug_script)
            sync_print(f"[DEBUG] Participant panel search:")
            sync_print(f"  - zmu-portal-dropdown-participant-list: {debug_info.get('byId', 'ERROR')}")
            sync_print(f"  - participants-section-container_wrapper: {debug_info.get('byClass', 0)} found")
            sync_print(f"  - Total LI items on page: {len(debug_info.get('liTexts', []))}")
            if debug_info.get('liTexts'):
                sync_print(f"  - LI texts: {debug_info['liTexts'][:5]}")
            if debug_info.get('divsWithText'):
                sync_print(f"  - Participant divs found: {len(debug_info['divsWithText'])}")
                for item in debug_info['divsWithText'][:3]:
                    sync_print(f"    * Class: {item['class']}, Text: {item['text']}")
        except Exception as e:
            sync_print(f"[DEBUG] Error running debug script: {e}")

        # Candidate selectors for participant name elements
        name_selectors = [
            '#zmu-portal-dropdown-participant-list li',  # Zoom web client structure
            '.participants-section-container_wrapper li',
            'div[class*="participant"] span',
            'div[class*="participant-name"]',
            '.participants-item__display-name',
            '.participants-item__name',
            '.participant-name',
            '.name',
            '.participants-list li',
            '.participants-list div',
            '[data-testid*="participant"]',
            '.zm-participant-name',
        ]

        names = []
        for sel in name_selectors:
            try:
                elems = driver.find_elements(By.CSS_SELECTOR, sel)
                for e in elems:
                    text = e.text.strip()
                    if text and text not in names and len(text) > 1:
                        # Filter out unwanted UI elements
                        import re
                        unwanted_patterns = [
                            'Unmute', 'start Video', 'Participants', 'chat', 'Reactions', 
                            'Share Screen', 'more', 'leave', 'Pleader', 'upgrade your browser',
                            'update your browder', 'Speaker', 'Gallery View', 'Participant \(',
                            'Mute', 'Turn off', 'Invite', 'Record', 'Security', 'Manage Participants'
                        ]
                        
                        is_unwanted = False
                        for pattern in unwanted_patterns:
                            if pattern.lower() in text.lower():
                                is_unwanted = True
                                break
                        
                        # Also filter out text with numbers in parentheses like "(3)"
                        if re.search(r'\(\s*\d+\s*\)', text):
                            is_unwanted = True
                            
                        if not is_unwanted:
                            names.append(text)
                if names:
                    sync_print(f"[DEBUG] Found names using selector: {sel}")
                    return names
            except Exception:
                pass
        
        return names

    sync_print(f"{name} sleep for {wait_time} seconds ...")
    
    participants_found = False
    last_participant_check = 0
    check_interval = 5  # Check for participants every 5 seconds
    
    try:
        elapsed = 0
        while not stop_event.is_set() and elapsed < wait_time:
            TimeNow = datetime.now().strftime('%H%M')
            if str(TimeNow) == str(Time):
                stop_event.set()
                break
            
            # Periodically check for participants
            if elapsed - last_participant_check >= check_interval:
                try:
                    participants = fetch_participants(driver)
                    if participants and not participants_found:
                        sync_print(f"Participants found ({len(participants)}): {participants}")
                        participants_found = True
                    elif participants and participants_found:
                        # Still print updates if new participants join
                        sync_print(f"Participants updated ({len(participants)}): {participants}")
                except Exception as e:
                    pass
                last_participant_check = elapsed
            
            time.sleep(1)
            elapsed += 1
    except KeyboardInterrupt:
        stop_event.set()
    finally:
        try:
            driver.quit()
        except Exception:
            pass
        sync_print(f"{name} ended!")


def main(headless_mode=False):
    wait_time = sec * 60
    workers = []
    for i in range(number):
        try:
            proxy = proxylist[i]
        except IndexError:
            proxy = None
        try:
            user = fake.name()
        except IndexError:
            break
        wk = threading.Thread(target=start, args=(
            f'[Thread{i}]', proxy, user, wait_time, headless_mode))
        workers.append(wk)
    for wk in workers:
        wk.start()
    for wk in workers:
        wk.join()


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Zoom joiner')
    parser.add_argument('number', nargs='?', type=int, help='Number of members to spawn')
    parser.add_argument('meetingcode', nargs='?', help='Meeting code')
    parser.add_argument('passcode', nargs='?', help='Meeting passcode')
    parser.add_argument('Time', nargs='?', help='End time HHMM')
    parser.add_argument('--headless', action='store_true', default=False, help='Run in headless mode (default: visible browser)')
    args = parser.parse_args()

    if args.number and args.meetingcode and args.passcode:
        number = args.number
        meetingcode = args.meetingcode
        passcode = args.passcode
        Time = args.Time or input("Enter Time for Ending (HHMM):")
    else:
        n = input("Enter Number of Members: ")
        number = int(n)
        meetingcode = input("Enter Meeting Code:")
        passcode = input("Enter Passcode:")
        Time = input("Enter Time for Ending (HHMM):")

    sec = 4
    
    def signal_handler(signum, frame):
        sync_print("\nCtrl+C detected. Shutting down gracefully...")
        stop_event.set()
    
    import signal
    signal.signal(signal.SIGINT, signal_handler)
    
    if important_request == "true":
        try:
            main(headless_mode=args.headless)
        except KeyboardInterrupt:
            stop_event.set()
        finally:
            try:
                if display:
                    display.stop()
            except Exception:
                pass
    else:
        print("Down")
