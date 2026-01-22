import os
from playwright.sync_api import sync_playwright

def test_add_new_row_id(page):
    # Load the local HTML file
    file_path = os.path.abspath("Index.html")
    page.goto(f"file://{file_path}")

    # Inject mock google.script.run
    page.add_init_script("""
        window.google = {
            script: {
                run: {
                    withSuccessHandler: function(callback) {
                        this.callback = callback;
                        return this;
                    },
                    withFailureHandler: function(callback) {
                        return this;
                    },
                    apiLogin: function(u, p) {
                        if (this.callback) this.callback({
                            success: true,
                            role: 'ADMIN',
                            name: 'Test Admin',
                            username: 'ADMIN'
                        });
                    },
                    getSystemConfig: function(role) {
                        if (this.callback) this.callback({
                            departments: { "DEP1": { label: "Department 1", color: "blue", icon: "fa-user" } },
                            staff: [{ name: "User1", dept: "DEP1" }],
                            directory: [],
                            specialModules: []
                        });
                    },
                    apiFetchStaffTrackerData: function(name) {
                        if (this.callback) this.callback({
                            success: true,
                            data: [],
                            history: [],
                            headers: ['FOLIO', 'CONCEPTO', 'ESTATUS']
                        });
                    },
                    apiUpdateTask: function(sheet, task) {
                        if (this.callback) this.callback({ success: true });
                    },
                    uploadFileToDrive: function() {},
                    apiFetchCascadeTree: function() {
                         if (this.callback) this.callback({ success: true, data: [] });
                    },
                    apiFetchDrafts: function() {
                         if (this.callback) this.callback({ success: true, data: [] });
                    }
                }
            }
        };
    """)

    # Reload to apply mock
    page.reload()

    # Login
    page.fill("input[placeholder='Usuario']", "ADMIN")
    page.fill("input[placeholder='Contraseña...']", "123")
    page.click("button:has-text('INICIAR SESIÓN')")

    # Wait for dashboard
    page.wait_for_selector("text=Dashboard", timeout=5000)

    # Click Department
    page.click("text=Department 1")

    # Click Staff Card
    page.click("text=User1")

    # Wait for Staff Tracker
    page.wait_for_selector("text=User1", timeout=5000)

    # Click "Agregar Fila" (Add Row)
    page.click("button[title='Agregar']")

    # Verify new row ID
    page.wait_for_timeout(1000) # Wait for Vue to render

    page.screenshot(path="verification/frontend_id_check.png")

    # Check FOLIO input value
    folio_input = page.locator("table.table-excel tbody tr").first.locator("td").nth(1).locator("input")
    folio_value = folio_input.input_value()
    print(f"Generated ID: {folio_value}")

    if not folio_value.startswith("VEN-"):
        raise Exception(f"ID does not start with VEN-: {folio_value}")

    # Check if it looks like a UUID (length > 15)
    if len(folio_value) < 15:
         print(f"Warning: ID seems short: {folio_value}")
         # If crypto.randomUUID is not available in this env, it might fall back to math.random
         # But in Playwright browser (Chromium), it should be available.

    if "NaN" in folio_value:
        raise Exception(f"ID contains NaN: {folio_value}")

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page()
    try:
        test_add_new_row_id(page)
    except Exception as e:
        print(f"Error: {e}")
        page.screenshot(path="verification/error_retry.png")
    finally:
        browser.close()
