import os
from playwright.sync_api import sync_playwright

def run(playwright):
    browser = playwright.chromium.launch(headless=True)
    page = browser.new_page()

    # Capture console logs
    page.on("console", lambda msg: print(f"Console: {msg.text}"))
    page.on("pageerror", lambda err: print(f"Page Error: {err}"))

    # Inject mock google.script.run
    page.add_init_script("""
    window.google = {
      script: {
        run: {
          withSuccessHandler: function(cb) {
            return {
              withFailureHandler: function(fcb) { return this; },
              apiLogin: function(u, p) {
                  setTimeout(() => cb({success:true, name:'Test User', role:'ADMIN'}), 100);
              },
              getSystemConfig: function(role) {
                  setTimeout(() => cb({
                     departments: { 'TEST': { label: 'Test Dept', color: 'red', icon: 'fa-user' } },
                     staff: [{name: 'TEST USER', dept: 'TEST'}],
                     directory: [],
                     specialModules: []
                  }), 100);
              },
              apiFetchStaffTrackerData: function(name) {
                  setTimeout(() => cb({ success:true, data: [], headers: ['ID', 'CONCEPTO'], history: [] }), 100);
              },
              apiUpdateTask: function(name, row) {
                  setTimeout(() => cb({ success:true }), 100);
              },
              apiFetchDrafts: function() {
                  setTimeout(() => cb({ success:true, data: [] }), 100);
              },
              apiFetchCascadeTree: function() {
                  setTimeout(() => cb({ success:true, data: [] }), 100);
              }
            };
          }
        }
      }
    };
    """)

    # Load the local HTML file
    cwd = os.getcwd()
    page.goto(f"file://{cwd}/Index.html")

    # Login
    try:
        page.fill("input[placeholder='Usuario']", "admin")
        page.fill("input[placeholder='Contraseña...']", "123")
        page.click("button:has-text('INICIAR SESIÓN')")

        # Wait for dashboard
        page.wait_for_selector(".dept-card", timeout=5000)

        # Click Test Dept
        page.click(".dept-card:has-text('Test Dept')")

        # Click Test User
        page.wait_for_selector(".staff-card", timeout=5000)
        page.click(".staff-card:has-text('TEST USER')")

        # Wait for table headers to ensure data is loaded
        print("Waiting for table...")
        page.wait_for_selector(".table-excel th", timeout=5000)

        # Click Add Row (Fila)
        page.wait_for_selector("button[title='Agregar']", timeout=5000)
        print("Clicking Add Row button...")
        page.click("button[title='Agregar']")

        # Wait for row
        page.wait_for_timeout(1000)

        # Check if row added
        rows = page.locator("tbody tr")
        print(f"Row count: {rows.count()}")

        id_inputs = page.locator("input.excel-input[readonly]")
        if id_inputs.count() > 0:
            id_value = id_inputs.first.input_value()
            print(f"Generated ID: {id_value}")

            if id_value.startswith("VEN-") and len(id_value) > 20:
                print("SUCCESS: ID format looks like VEN-UUID")
            else:
                print(f"FAILURE: ID format incorrect: {id_value}")
        else:
             print("FAILURE: No ID input found")

        page.screenshot(path="verification/verification_uuid.png")
    except Exception as e:
        print(f"Error: {e}")
        page.screenshot(path="verification/error.png")

    browser.close()

if __name__ == "__main__":
    with sync_playwright() as p:
        run(p)
