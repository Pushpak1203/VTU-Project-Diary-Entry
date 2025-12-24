from seleniumbase import SB
import pandas as pd
import os
from datetime import datetime
import time


time.sleep(5)

user_mail = "YOUR MAIL"
user_passwd = "PASSWORD"

with SB(uc=True) as sb:
    login_url = "https://vtu.internyet.in/sign-in"
    sb.open(login_url)
    
    sb.type('input[autocomplete="email"]', user_mail)
    sb.type('input[name="password"]', user_passwd)
    sb.click('button[type="submit"]')
    
    sb.sleep(5) 
    
    diary_url = "https://vtu.internyet.in/dashboard/student/project-diary"
    sb.open(diary_url)
    sb.sleep(3)

    print("Navigated to Project Diary Entries page.")

    excel_path = "Complete_Project_Diary.xlsx"
    if not os.path.exists(excel_path):
        print(f"Error: Excel file not found at {excel_path}")
    else:
        try:
            df = pd.read_excel(excel_path)
            print(f"Loaded {len(df)} rows from Excel.")
            
            for index, row in df.iterrows():
                print(f"Processing row {index + 1}...")
                
                date_str = str(row['Date']).strip() 
                work_summary = str(row['Work Summary'])
                hours_worked = str(row['Hours Worked'])
                learnings = str(row['Learning Outcomes'])
                blockers = str(row['Blockers/Risks'])
                
                try:
                    target_date = datetime.strptime(date_str, "%B %d, %Y")
                except ValueError:
                    target_date = pd.to_datetime(date_str).to_pydatetime()
                
                target_day_str = target_date.strftime("%Y-%m-%d")
                print(f"Target Date: {target_day_str}")

                print("Selecting Project...")
                sb.wait_for_element_visible('select[name="project_id"]', timeout=20)
                project_script = """
                (function() {
                    const selectEl = document.querySelector('select[name="project_id"]');
                    if (!selectEl) {
                        return;
                    }

                    const firstOption = selectEl.options[0];
                    if (!firstOption) {
                        return;
                    }

                    const optionValue = firstOption.value || firstOption.textContent.trim();
                    selectEl.value = optionValue;

                    selectEl.dispatchEvent(new Event("input", { bubbles: true }));
                    selectEl.dispatchEvent(new Event("change", { bubbles: true }));
                })();
                """
                sb.execute_script(project_script)
                sb.sleep(1)

                # print("Selecting Date...")
                
                js_month_val = target_date.month - 1
                js_year_val = target_date.year
                js_day_text = str(target_date.day) 
                
                date_picker_script = f"""
                (function() {{
                    const TARGET_YEAR = "{js_year_val}";
                    const TARGET_MONTH_VALUE = "{js_month_val}";
                    const TARGET_DAY_TEXT = "{js_day_text}";

                    const existingDialog = document.querySelector('[role="dialog"]');
                    if (!existingDialog) {{
                        const trigger = document.querySelector('button[aria-haspopup="dialog"]');
                        if (trigger) {{
                            trigger.click();
                        }} else {{
                            return false;
                        }}
                    }}

                    function waitForMonthAndProceed() {{
                        const monthSelect = document.querySelector('select.rdp-months_dropdown');
                        
                        if (!monthSelect) {{
                            requestAnimationFrame(waitForMonthAndProceed);
                            return;
                        }}

                        const container = monthSelect.closest('[role="dialog"]') ||
                                          monthSelect.closest('.rdp-popover') ||
                                          monthSelect.parentElement ||
                                          document;

                        try {{
                            monthSelect.value = TARGET_MONTH_VALUE;
                            monthSelect.dispatchEvent(new Event("change", {{ bubbles: true }}));
                        }} catch (e) {{}}

                        const yearSelect = container.querySelector('select.rdp-years_dropdown') || document.querySelector('select.rdp-years_dropdown');
                        if (yearSelect) {{
                            try {{
                                yearSelect.value = TARGET_YEAR;
                                yearSelect.dispatchEvent(new Event("change", {{ bubbles: true }}));
                            }} catch (e) {{}}
                        }}

                        let attempts = 0;
                        const maxAttempts = 20;
                        
                        const intervalId = setInterval(() => {{
                            attempts++;
                            const container = document.querySelector('[role="dialog"]') || 
                                              document.querySelector('.rdp-popover') || 
                                              document;
                            
                            const candidates = Array.from(container.querySelectorAll('button'));
                            
                            const dayBtn = candidates.find(b => {{
                                const text = b.textContent ? b.textContent.trim() : "";
                                const isOutside = b.classList.contains("rdp-day_outside") || b.className.includes("outside");
                                const isVisible = b.offsetParent !== null;
                                
                                return text === TARGET_DAY_TEXT && !b.disabled && isVisible && !isOutside;
                            }});

                            if (dayBtn) {{
                                dayBtn.click();
                                clearInterval(intervalId);
                            }} else {{
                                if (attempts >= maxAttempts) {{
                                    clearInterval(intervalId);
                                    
                                    const fallback = Array.from(document.querySelectorAll('button'))
                                        .find(b => {{
                                            const text = b.textContent ? b.textContent.trim() : "";
                                            const isOutside = b.classList.contains("rdp-day_outside") || b.className.includes("outside");
                                            const isVisible = b.offsetParent !== null;
                                            return text === TARGET_DAY_TEXT && !b.disabled && isVisible && !isOutside;
                                        }});
                                    if (fallback) {{
                                        fallback.click();
                                    }}
                                }}
                            }}
                        }}, 500); 
                    }}

                    waitForMonthAndProceed();
                }})();
                """
                sb.execute_script(date_picker_script)
                sb.sleep(3) 

                # Retry logic for clicking Continue
                max_retries = 3
                for attempt in range(max_retries):
                    try:
                        # print(f"Clicking Continue (Attempt {attempt + 1})...")
                        sb.wait_for_element_visible("//button[contains(text(), 'Continue')]", by="xpath", timeout=5)
                        sb.click("//button[contains(text(), 'Continue')]", by="xpath")
                        
                        # Check if successful by waiting for the next element
                        sb.wait_for_element_visible('textarea[name="description"]', timeout=5)
                        break # Success, exit loop
                    except Exception:
                        if attempt == max_retries - 1:
                            print("Failed to navigate to description page after retries.")
                            raise
                        print("Retrying Continue click...")
                        sb.sleep(2)

                sb.type('textarea[name="description"]', work_summary)
                sb.type('input[placeholder="e.g. 6.5"]', hours_worked)
                sb.type('textarea[name="learnings"]', learnings)
                sb.type('textarea[name="blockers"]', blockers)

                # print("Entering Skills...")
                skill_script = """
                (function selectSkillReact() {
                    const input = document.querySelector('.react-select__input');
                    if (!input) {
                        return;
                    }

                    input.focus();

                    const nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, "value").set;
                    nativeInputValueSetter.call(input, "React.js");

                    input.dispatchEvent(new Event("input", { bubbles: true }));

                    ["keydown", "keyup"].forEach(type => {
                        input.dispatchEvent(new KeyboardEvent(type, {
                            key: "Enter",
                            code: "Enter",
                            keyCode: 13,
                            which: 13,
                            bubbles: true,
                        }));
                    });
                })();
                """
                sb.execute_script(skill_script)
                sb.sleep(1)

                # print("Clicking Save...")
                try:
                    sb.execute_script("document.evaluate(\"//button[contains(text(), 'Save')]\", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.click()")
                except:
                    sb.click('button[type="submit"]')
                
                # print("Waiting for save to complete...")
                sb.sleep(2) 
                sb.sleep(1)
                
                # print("Clicking Create Next...")
                try:
                    sb.wait_for_element_visible("//a[contains(text(), 'Create') and contains(@href, '/dashboard/student/project-diary')]", by="xpath", timeout=10)
                    sb.click("//a[contains(text(), 'Create') and contains(@href, '/dashboard/student/project-diary')]", by="xpath")
                except Exception as e:
                    sb.open("https://vtu.internyet.in/dashboard/student/project-diary")

                sb.sleep(2)
                
        except Exception as e:
            print(f"An error occurred during processing: {e}")
