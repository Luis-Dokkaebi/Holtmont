from playwright.sync_api import sync_playwright
import os

def run(playwright):
    browser = playwright.chromium.launch()
    page = browser.new_page()

    # Load the local HTML file
    cwd = os.getcwd()
    page.goto(f"file://{cwd}/Index.html")

    # Inject mock google.script.run
    page.evaluate("""
        window.google = {
            script: {
                run: {
                    withSuccessHandler: function(handler) {
                        return {
                            withFailureHandler: function(errHandler) {
                                return this;
                            },
                            apiLogin: function(u, p) {
                                handler({success: true, role: 'ADMIN', name: 'TestUser', username: 'TEST'});
                            },
                            getSystemConfig: function(role) {
                                handler({
                                    departments: { 'DEP1': { label: 'Dept 1', icon: 'fa-home', color: 'red' } },
                                    staff: [{name: 'User1', dept: 'DEP1'}],
                                    directory: [],
                                    specialModules: []
                                });
                            },
                            apiFetchDrafts: function() {
                                handler({success: true, data: []});
                            },
                            apiFetchCascadeTree: function() {
                                handler({success: true, data: []});
                            },
                            apiFetchStaffTrackerData: function(name) {
                                handler({
                                    success: true,
                                    data: [],
                                    history: [],
                                    headers: ['ID', 'CONCEPTO', 'FECHA', 'VENDEDOR']
                                });
                            },
                            apiFetchSalesHistory: function() {
                                handler({success: true, data: {}});
                            }
                        };
                    },
                    apiSyncDrafts: function() {},
                    apiFetchSalesHistory: function() {}
                }
            }
        };
    """)

    # Login
    page.fill('input[placeholder="Usuario"]', "admin")
    page.fill('input[placeholder="Contraseña..."]', "admin")
    page.click('button:has-text("INICIAR SESIÓN")')

    # Wait for dashboard
    page.wait_for_selector('.dept-card', state='visible')

    # Go to a staff tracker
    page.click('.dept-card') # Click first dept
    page.wait_for_selector('.staff-card')
    page.click('.staff-card') # Click first staff

    # Wait for tracker view
    page.wait_for_selector('.table-excel')

    # Add new row
    page.click('button[title="Agregar"]')

    # Screenshot
    page.screenshot(path="verification/verification.png")

    # Extract the ID
    id_val = page.evaluate("""
        () => {
            const inputs = Array.from(document.querySelectorAll('input'));
            const idInput = inputs.find(i => i.value && i.value.startsWith('VEN-'));
            return idInput ? idInput.value : "NOT_FOUND";
        }
    """)

    print(f"Generated ID: {id_val}")

    browser.close()

if __name__ == "__main__":
    with sync_playwright() as playwright:
        run(playwright)
