import asyncio
import re
from typing import Callable, List, Tuple
from playwright.async_api import async_playwright, Page

BASE_URL = "https://sipp.petroil.com.mx/login.html"

# ──────────────────────────────────────────────────────────
# JavaScript helpers that talk directly to AngularJS scopes
# ──────────────────────────────────────────────────────────
_JS_SET_EMPRESA = """() => {
    const sel = document.querySelector("select[ng-model='id_Empresa']");
    if (!sel) return false;
    const opt = Array.from(sel.options).find(o =>
        o.text.trim().toUpperCase().startsWith('PETROPLAZAS -')
    );
    if (!opt) return false;
    // Set native value and fire change so Angular + chosen both react
    sel.value = opt.value;
    sel.dispatchEvent(new Event('change', { bubbles: true }));
    // Also trigger via Angular scope to be safe
    try {
        const scope = angular.element(sel).scope();
        scope.$apply(() => { scope.id_Empresa = opt.value; });
    } catch(e) {}
    // Tell chosen to refresh its UI
    if (typeof $ !== 'undefined') { $(sel).trigger('chosen:updated'); }
    return true;
}"""

_JS_SET_SUCURSAL = """() => {
    const sel = document.querySelector("select[ng-model='id_Sucursal']");
    if (!sel) return false;
    const opt = Array.from(sel.options).find(o =>
        o.text.toUpperCase().includes('CORPORATIVO')
    );
    if (!opt) return false;
    sel.value = opt.value;
    sel.dispatchEvent(new Event('change', { bubbles: true }));
    try {
        const scope = angular.element(sel).scope();
        scope.$apply(() => { scope.id_Sucursal = opt.value; });
    } catch(e) {}
    if (typeof $ !== 'undefined') { $(sel).trigger('chosen:updated'); }
    return true;
}"""

_JS_SUCURSAL_LOADED = """() => {
    const sel = document.querySelector("select[ng-model='id_Sucursal']");
    return Boolean(sel && sel.options.length > 1);
}"""

_JS_SET_ESTATUS_VACIO = """() => {
    const sel = document.querySelector("select[ng-model='filtro.id_Estatus']");
    if (!sel) return;
    const scope = angular.element(sel).scope();
    scope.$apply(() => { scope.filtro.id_Estatus = ''; });
}"""

_JS_GRID_ROW_COUNT = """(gridAttr) => {
    const grid = document.querySelector(`[ng-grid="${gridAttr}"]`);
    return grid ? grid.querySelectorAll('.ngRow').length : 0;
}"""


class RPAAutomation:
    def __init__(
        self,
        username: str,
        password: str,
        headless: bool = False,
        log_fn: Callable = print,
        cancel_fn: Callable = lambda: False,
    ):
        self.username = username
        self.password = password
        self.headless = headless
        self.log = log_fn
        self.should_cancel = cancel_fn
        self.skipped: List[str] = []
        self.not_found: List[str] = []

    # ──────────────────────────────────────────────────────
    # Public entry point
    # ──────────────────────────────────────────────────────
    async def run(
        self,
        folio_rows: List[Tuple[int, str]],
        on_progress: Callable = None,
    ) -> List[Tuple]:
        """
        Process every (row_num, folio) pair and return list of
        (row_num, cc, observaciones, subtotal, descuento, iva, gastos_envio, total_oc).
        """
        results: List[Tuple] = []

        async with async_playwright() as p:
            browser = await p.chromium.launch(
                headless=self.headless,
                slow_mo=80,
                args=["--start-maximized"],
            )
            context = await browser.new_context(
                viewport={"width": 1440, "height": 900},
                locale="es-MX",
            )
            page = await context.new_page()

            # Dismiss any browser dialogs automatically
            page.on("dialog", lambda d: asyncio.ensure_future(d.accept()))

            try:
                await self._login(page)
                await self._configure_session(page)
                await self._navigate_to_recepcion(page)

                processed = 0
                errors = 0
                seen: set = set()

                for row_num, folio in folio_rows:
                    if self.should_cancel():
                        self.log("Proceso cancelado por el usuario.", "warn")
                        break

                    folio = str(folio).strip()

                    # Duplicate guard
                    if folio in seen:
                        self.log(f"Folio duplicado omitido: {folio}", "warn")
                        self.skipped.append(folio)
                        continue
                    seen.add(folio)

                    if on_progress:
                        on_progress(processed, errors, folio)

                    try:
                        self.log(f"Procesando folio: {folio}", "info")
                        cc, obs, subtotal, descuento, iva, gastos_envio, total_oc, cuentas_contables = \
                            await self._process_folio(page, folio)
                        results.append((row_num, cc, obs, subtotal, descuento, iva, gastos_envio, total_oc, cuentas_contables))

                        if cc:
                            self.log(f"  CC: {cc}", "ok")
                        else:
                            self.log(f"  Sin datos de OC (folio: {folio})", "warn")

                        processed += 1

                    except Exception as exc:
                        self.log(f"  Error en folio {folio}: {exc}", "error")
                        results.append((row_num, "", "", "", "", "", "", "", []))
                        errors += 1
                        await self._recover_page(page)

                    if on_progress:
                        on_progress(processed, errors, folio)

            finally:
                await browser.close()

        return results

    # ──────────────────────────────────────────────────────
    # Step 1 — Login
    # ──────────────────────────────────────────────────────
    async def _login(self, page: Page):
        self.log("Abriendo página de login...", "info")
        await page.goto(BASE_URL, wait_until="networkidle", timeout=30_000)
        await page.wait_for_selector("#btnLogin", state="visible", timeout=15_000)
        await page.wait_for_timeout(400)

        self.log("Ingresando credenciales...", "info")
        await page.fill("#nb_Usuario", self.username)
        await page.fill("input[ng-model='de_password']", self.password)
        await page.wait_for_timeout(300)
        await page.click("#btnLogin")

        # Wait until we leave the login page
        await page.wait_for_function(
            "() => !window.location.href.includes('login.html')",
            timeout=30_000,
        )
        self.log("Login exitoso.", "ok")

    # ──────────────────────────────────────────────────────
    # Step 2 — Select company & branch via Chosen UI clicks
    # ──────────────────────────────────────────────────────
    async def _configure_session(self, page: Page):
        self.log("Configurando sesión...", "info")

        # Wait for the page and chosen to fully initialise
        await page.wait_for_selector(".chosen-container", state="visible", timeout=20_000)
        await page.wait_for_timeout(800)

        # Close password-update modal if it appears
        pwd_modal = page.locator("#divBloqueo_modalActualizarContrasena")
        if await pwd_modal.is_visible():
            self.log("Cerrando modal de contraseña predeterminada...", "warn")
            await page.locator(
                "#divBloqueo_modalActualizarContrasena .btn-cerrar25p"
            ).click()
            await page.wait_for_timeout(500)

        # ── Empresa: use Chosen UI so Angular sees a real user interaction ──
        # The Empresa chosen-container is the one whose underlying select has ng-model='id_Empresa'
        await self._chosen_select(page, "id_Empresa", "PETROPLAZAS -")
        self.log("Empresa seleccionada: PETROPLAZAS", "ok")
        await page.wait_for_timeout(1_500)

        # Wait for Sucursal options to load (server round-trip after empresa change)
        self.log("Esperando carga de sucursales...", "info")
        await page.wait_for_function(_JS_SUCURSAL_LOADED, timeout=15_000)
        await page.wait_for_timeout(500)

        # ── Sucursal ──
        await self._chosen_select(page, "id_Sucursal", "CORPORATIVO")
        self.log("Sucursal seleccionada: PETROPLAZAS CORPORATIVO", "ok")
        await page.wait_for_timeout(600)

        # Save session
        await page.click("button[ng-click='Guardar()']")
        await page.wait_for_timeout(2_500)
        self.log("Sesión guardada.", "ok")

    async def _chosen_select(self, page: Page, ng_model: str, text_filter: str):
        """
        Interact with a chosen-enhanced <select> by clicking through its UI.
        Finds the chosen container associated with the select that has the given
        ng-model, opens it, types to filter, and clicks the matching option.
        """
        # Find the chosen container via JS (it's inserted right after the hidden select)
        container_id = await page.evaluate(f"""() => {{
            const sel = document.querySelector("select[ng-model='{ng_model}']");
            if (!sel) return null;
            // chosen inserts a sibling div.chosen-container after the select
            let node = sel.nextElementSibling;
            while (node) {{
                if (node.classList && node.classList.contains('chosen-container')) {{
                    // Give it a temp id so Playwright can target it
                    if (!node.id) node.id = 'rpa_chosen_{ng_model}';
                    return node.id;
                }}
                node = node.nextElementSibling;
            }}
            return null;
        }}""")

        if not container_id:
            raise RuntimeError(f"No se encontró chosen-container para ng-model='{ng_model}'")

        container = page.locator(f"#{container_id}")

        # Click to open the dropdown
        await container.locator("a.chosen-single").click()
        await page.wait_for_timeout(300)

        # Type the filter text into the search box
        search_input = container.locator(".chosen-search input")
        await search_input.fill(text_filter)
        await page.wait_for_timeout(400)

        # Click the first visible matching result
        result = container.locator(
            f".chosen-results li.active-result:has-text('{text_filter}')"
        ).first
        await result.wait_for(state="visible", timeout=5_000)
        await result.click()
        await page.wait_for_timeout(300)

    # ──────────────────────────────────────────────────────
    # Step 3 — Navigate to Recepción de Facturas
    # ──────────────────────────────────────────────────────
    async def _navigate_to_recepcion(self, page: Page):
        self.log("Navegando a Recepción de Facturas...", "info")
        base = page.url.split("#")[0]
        await page.goto(
            f"{base}#/RecepcionFacturas",
            wait_until="networkidle",
            timeout=30_000,
        )
        await page.wait_for_selector(
            "input[ng-model='filtro.nu_foliodocumento']",
            timeout=20_000,
        )
        # Pre-set Estatus to "Seleccionar" once; we keep it that way throughout
        await page.evaluate(_JS_SET_ESTATUS_VACIO)
        self.log("Página Recepción de Facturas lista.", "ok")

    # ──────────────────────────────────────────────────────
    # Step 4 — Process a single folio
    # ──────────────────────────────────────────────────────
    async def _process_folio(self, page: Page, folio: str) -> Tuple:
        # Safety: dismiss any lingering <red-alert> from a previous "no encontrado"
        await self._dismiss_red_alert(page)

        # Fill folio field (fill() clears existing content automatically)
        folio_input = page.locator("input[ng-model='filtro.nu_foliodocumento']")
        await folio_input.fill("")
        await folio_input.fill(folio)
        await page.evaluate(_JS_SET_ESTATUS_VACIO)
        await page.wait_for_timeout(200)

        # Execute search
        await page.click("button[ng-click='buscar()']")
        await page.wait_for_timeout(2_000)

        # Check if SIPP showed a "not found" alert and dismiss it
        _EMPTY8 = ("", "", "", "", "", "", "", [])

        if await self._dismiss_red_alert(page):
            self.log(f"  Folio {folio}: no encontrado en SIPP.", "warn")
            self.not_found.append(folio)
            return _EMPTY8

        await page.wait_for_timeout(500)

        # Verify rows loaded in the main list grid
        row_count = await page.evaluate(_JS_GRID_ROW_COUNT, "listadoGrid")
        if row_count == 0:
            self.log(f"  Folio {folio}: sin resultados en SIPP.", "warn")
            self.not_found.append(folio)
            return _EMPTY8

        # ── Click "Visualizar Detalle" on first row ──
        first_row = page.locator("[ng-grid='listadoGrid'] .ngRow").first
        detail_btn = await self._find_btn(
            first_row,
            ["[title='Visualizar Detalle']", "[title*='Detalle']"],
            fallback_index=0,
        )
        await detail_btn.click()

        # Wait for Visualizar Factura modal
        await page.wait_for_selector(
            "#content_modalVisualizar", state="visible", timeout=12_000
        )
        await page.wait_for_timeout(1_000)

        # ── Inside the modal, extract OC data + Cuenta Contable ──
        result = await self._extract_from_visualizar_modal(page)

        # Close Visualizar modal
        await self._close_modal(page, "content_modalVisualizar")
        return result

    # ──────────────────────────────────────────────────────
    # Extract CC + Observaciones + Cuenta Contable from "Visualizar Factura" modal
    # ──────────────────────────────────────────────────────
    async def _extract_from_visualizar_modal(
        self, page: Page
    ) -> Tuple:
        modal = page.locator("#content_modalVisualizar")

        # Give the modal's Angular controller time to fetch Servicios via API
        await page.wait_for_timeout(2_500)

        # Strategy 1: wait for .ngRow elements inside movimientosDetalleGrid
        cuentas_contables: list = []
        doc_btn = None
        try:
            await page.wait_for_function(
                """() => {
                    const grid = document.querySelector(
                        '#content_modalVisualizar [ng-grid="movimientosDetalleGrid"]'
                    );
                    return grid !== null && grid.querySelectorAll('.ngRow').length > 0;
                }""",
                timeout=12_000,
            )
            svc_rows = modal.locator("[ng-grid='movimientosDetalleGrid'] .ngRow")
            count = await svc_rows.count()
            self.log(f"  Servicios: {count} fila(s) en grid.", "info")

            # Extraer TODAS las cuentas contables — el modal ya está completamente cargado
            cuentas_contables = await self._extract_cuentas_contables(page)
            if cuentas_contables:
                self.log(f"  Cuentas Contables ({len(cuentas_contables)}): {', '.join(cuentas_contables)}", "info")

            if count > 0:
                doc_btn = await self._find_btn(
                    svc_rows.first,
                    ["[title='Visualizar Documento']", "[title*='Documento']"],
                    fallback_index=1,
                )
        except Exception:
            pass

        # Strategy 2: fallback — search for the button anywhere inside the modal
        if doc_btn is None:
            self.log("  Buscando botón Visualizar Documento por fallback...", "warn")
            fallback = modal.locator(
                "[title='Visualizar Documento'], [title*='Previsualizar Documento'],"
                " .btn-icon25p:nth-child(2)"
            ).first
            if await fallback.count() > 0:
                doc_btn = fallback
            else:
                if not cuentas_contables:
                    cuentas_contables = await self._extract_cuentas_contables(page)
                self.log("  Sin sección Servicios para este folio.", "warn")
                return "", "", "", "", "", "", "", cuentas_contables

        await doc_btn.click()

        # Wait for OC document modal
        await page.wait_for_selector(
            "#content_modalDocOC", state="visible", timeout=12_000
        )
        await page.wait_for_timeout(1_000)

        cc, obs, subtotal, descuento, iva, gastos_envio, total_oc = \
            await self._extract_from_doc_modal(page)

        # Close OC document modal
        await self._close_modal(page, "content_modalDocOC")
        return cc, obs, subtotal, descuento, iva, gastos_envio, total_oc, cuentas_contables

    # ──────────────────────────────────────────────────────
    # Extraer TODAS las Cuentas Contables del modal Visualizar Detalle
    # ──────────────────────────────────────────────────────
    async def _extract_cuentas_contables(self, page: Page) -> list:
        """
        Extrae todos los códigos de cuenta del apartado 'Cuentas Contables'
        en el modal #content_modalVisualizar. Retorna lista de strings (dígitos sin guiones).
        """
        try:
            result = await page.evaluate("""() => {
                const modal = document.querySelector('#content_modalVisualizar');
                if (!modal) return [];
                const results = [];

                // --- Estrategia 1: ng-grid que no sea movimiento ni detalle ---
                const grids = modal.querySelectorAll('[ng-grid]');
                for (const grid of grids) {
                    const attr = (grid.getAttribute('ng-grid') || '').toLowerCase();
                    if (attr.includes('movimiento') || attr.includes('detalle')) continue;

                    const rows = grid.querySelectorAll('.ngRow');
                    if (rows.length === 0) continue;

                    for (const row of rows) {
                        let found = '';
                        for (const inp of row.querySelectorAll('input')) {
                            const raw = (inp.value || '').trim();
                            if (!raw) continue;
                            const stripped = raw.replace(/-/g, '');
                            if (/^\\d{8,13}$/.test(stripped)) { found = stripped; break; }
                            // Cuenta cuadre: "CUADRE" al quitar guiones/underscores
                            if (raw.replace(/[-_\\s]/g, '').toUpperCase().includes('CUADRE')) {
                                found = raw; break;
                            }
                        }
                        if (!found) {
                            for (const cell of row.querySelectorAll('.ngCell')) {
                                const raw = cell.textContent.trim();
                                if (!raw) continue;
                                const stripped = raw.replace(/-/g, '');
                                if (/^\\d{8,13}$/.test(stripped)) { found = stripped; break; }
                                if (raw.replace(/[-_\\s]/g, '').toUpperCase().includes('CUADRE')) {
                                    found = raw; break;
                                }
                            }
                        }
                        if (found) results.push(found);
                    }
                    if (results.length > 0) return results;
                }

                // --- Estrategia 2: tabla HTML con header "Cuenta" ---
                for (const table of modal.querySelectorAll('table')) {
                    const headers = Array.from(
                        table.querySelectorAll('thead th, tr:first-child th')
                    ).map(h => h.textContent.trim());
                    const cuentaIdx = headers.findIndex(h => h === 'Cuenta');
                    if (cuentaIdx < 0) continue;

                    const dataRows = table.querySelectorAll('tbody tr');
                    for (const row of dataRows) {
                        const cells = row.querySelectorAll('td');
                        if (cells.length <= cuentaIdx) continue;
                        const cell = cells[cuentaIdx];
                        const inp = cell.querySelector('input');
                        const raw = inp
                            ? (inp.value || '').trim()
                            : cell.textContent.trim();
                        if (!raw) continue;
                        const stripped = raw.replace(/-/g, '');
                        if (/^\\d{8,13}$/.test(stripped)) { results.push(stripped); continue; }
                        if (raw.replace(/[-_\\s]/g, '').toUpperCase().includes('CUADRE')) {
                            results.push(raw); continue;
                        }
                    }
                    if (results.length > 0) return results;
                }

                return results;
            }""")
            return result if isinstance(result, list) else []
        except Exception:
            return []

    # ──────────────────────────────────────────────────────
    # Parse CC and Observaciones OC from the OC viewer modal
    # ──────────────────────────────────────────────────────
    async def _extract_from_doc_modal(self, page: Page) -> Tuple:
        """Returns (cc, obs, subtotal, descuento, iva, gastos_envio, total_oc)."""
        _EMPTY = ("", "", "", "", "", "", "")
        content_div = page.locator("#modal-bodymodalDocOC .ng-binding").first

        try:
            html_content = await content_div.inner_html(timeout=8_000)
        except Exception:
            self.log("  No se pudo leer el contenido del documento OC.", "warn")
            return _EMPTY

        if not html_content.strip():
            self.log("  Documento OC vacío.", "warn")
            return _EMPTY

        from bs4 import BeautifulSoup

        soup = BeautifulSoup(html_content, "lxml")
        cc  = self._extract_cc(soup)
        obs = self._extract_observaciones(soup)
        fin = self._extract_financials(soup)

        # Log financials so the user can verify extraction
        if any(fin.values()):
            self.log(
                f"  Financieros OC — Sub: {fin['subtotal']}  "
                f"Desc: {fin['descuento']}  IVA: {fin['iva']}  "
                f"Env: {fin['gastos_envio']}  Total: {fin['total']}",
                "info",
            )
        else:
            self.log("  Financieros OC: no detectados.", "warn")

        return cc, obs, fin["subtotal"], fin["descuento"], fin["iva"], fin["gastos_envio"], fin["total"]

    # ──────────────────────────────────────────────────────
    # HTML parsing helpers
    # ──────────────────────────────────────────────────────
    @staticmethod
    def _extract_financials(soup) -> dict:
        """Extract SUBTOTAL, DESCUENTO, IVA, GASTOS DE ENVÍO and TOTAL from OC doc."""
        text = soup.get_text("\n")
        # Amount pattern: optional $ + digits/commas/dots
        _AMT = r"\$?\s*([\d,]+\.?\d*)"
        patterns = {
            "subtotal":     rf"\bSUBTOTAL\s*:\s*{_AMT}",
            "descuento":    rf"\bDESCUENTO\s*:\s*{_AMT}",
            "iva":          rf"\bIVA\s*\([\d.]+\s*%\)\s*:\s*{_AMT}",
            "gastos_envio": rf"\bGASTOS\s+DE\s+ENV[IÍ]O\s*:\s*{_AMT}",
            # Negative lookbehind so "SUBTOTAL:" doesn't match "TOTAL:"
            "total":        rf"(?<!SUB)TOTAL\s*:\s*{_AMT}",
        }
        result = {}
        for key, pat in patterns.items():
            m = re.search(pat, text, re.IGNORECASE)
            result[key] = m.group(1).strip() if m else ""
        return result

    @staticmethod
    def _extract_cc(soup) -> str:
        """Find the value under the 'CC' column in the OC items table."""
        for table in soup.find_all("table"):
            header_row = table.find("tr")
            if not header_row:
                continue
            headers = [
                cell.get_text(strip=True)
                for cell in header_row.find_all(["th", "td"])
            ]
            if "CC" not in headers:
                continue
            idx = headers.index("CC")
            for tr in table.find_all("tr")[1:]:
                cells = tr.find_all("td")
                if len(cells) > idx:
                    val = cells[idx].get_text(strip=True)
                    if val:
                        return val

        # Fallback: scan plain text for "CC\n<value>" pattern
        lines = [ln.strip() for ln in soup.get_text("\n").splitlines() if ln.strip()]
        skip_words = {"CC", "Estatus", "Surtido", "Sub Total", "Grupo CC", "Insumo"}
        for i, ln in enumerate(lines):
            if ln == "CC" and i + 1 < len(lines):
                candidate = lines[i + 1]
                if candidate and candidate not in skip_words:
                    return candidate
        return ""

    @staticmethod
    def _extract_observaciones(soup) -> str:
        """Find text that follows 'Observaciones OC:' label in the OC document."""
        full_text = soup.get_text("\n")

        # "Observaciones OC" appears near the end of the doc, followed by "Autoriz" or end.
        # NOTE: Moneda/Importe appear BEFORE Observaciones, not after — don't use as stoppers.
        # Omit the accent from Autorizó to avoid encoding edge cases.
        match = re.search(
            r"Observaciones\s+OC\s*:?\s*(.+?)(?=\r?\nAutoriz|\r?\nDescargar\s+Aqu|\Z)",
            full_text,
            re.IGNORECASE | re.DOTALL,
        )
        if match:
            result = match.group(1).strip()
            if result:
                return result

        # Fallback A: "Observaciones" and "OC" may land on separate lines due to table cells.
        match2 = re.search(
            r"Observaciones\r?\n\s*OC\s*:?\s*(.+?)(?=\r?\nAutoriz|\r?\nDescargar|\Z)",
            full_text,
            re.IGNORECASE | re.DOTALL,
        )
        if match2:
            result2 = match2.group(1).strip()
            if result2:
                return result2

        # Fallback B: scan soup nodes for a node whose text starts with "Observaciones OC"
        for node in soup.find_all(string=re.compile(r"Observaciones", re.I)):
            container_text = (
                node.parent.get_text(" ", strip=True) if node.parent else ""
            )
            m3 = re.search(
                r"Observaciones\s*(?:OC)?\s*:?\s*(.+)", container_text, re.I | re.DOTALL
            )
            if m3:
                return m3.group(1).strip()

        # Fallback C: line-by-line scan
        lines = [ln.strip() for ln in full_text.splitlines()]
        for i, ln in enumerate(lines):
            if re.fullmatch(r"Observaciones\s+OC\s*:?", ln, re.I):
                # The value is on the next non-empty line
                for j in range(i + 1, min(i + 4, len(lines))):
                    if lines[j]:
                        return lines[j]
            m4 = re.match(r"Observaciones\s+OC\s*:\s*(.+)", ln, re.I)
            if m4:
                return m4.group(1).strip()

        return ""

    # ──────────────────────────────────────────────────────
    # Utility: find a button by title selectors or by index
    # ──────────────────────────────────────────────────────
    @staticmethod
    async def _find_btn(container, selectors: List[str], fallback_index: int = 0):
        for sel in selectors:
            loc = container.locator(sel)
            if await loc.count() > 0:
                return loc.first
        # Fallback: nth button/link in the container
        all_btns = container.locator("button, a[ng-click], a.btn")
        count = await all_btns.count()
        idx = min(fallback_index, max(0, count - 1))
        return all_btns.nth(idx)

    # ──────────────────────────────────────────────────────
    # Dismiss <red-alert> "no encontrado" overlay
    # ──────────────────────────────────────────────────────
    async def _dismiss_red_alert(self, page: Page) -> bool:
        """
        Click "Aceptar" inside a <red-alert> overlay if one is visible.
        Returns True if an alert was found and dismissed.
        The overlay's <td align="center"> intercepts pointer events when open.
        """
        try:
            alert = page.locator("red-alert")
            if await alert.count() == 0 or not await alert.is_visible():
                return False
            # Try common dismiss targets inside the alert (Aceptar button or the td itself)
            for sel in [
                "red-alert input[type='button']",
                "red-alert button",
                "red-alert .btn",
                "red-alert td[align='center']",
                "red-alert td",
            ]:
                btn = page.locator(sel).first
                if await btn.count() > 0 and await btn.is_visible():
                    await btn.click(timeout=4_000)
                    await page.wait_for_timeout(600)
                    self.log("  Alerta de SIPP cerrada (folio no encontrado).", "warn")
                    return True
        except Exception:
            pass
        return False

    # ──────────────────────────────────────────────────────
    # Modal close helpers
    # ──────────────────────────────────────────────────────
    @staticmethod
    async def _close_modal(page: Page, modal_id: str):
        try:
            btn = page.locator(f"#{modal_id} .btn-cerrar25p")
            if await btn.is_visible():
                await btn.click()
                await page.wait_for_timeout(500)
        except Exception:
            pass

    async def _recover_page(self, page: Page):
        """
        Best-effort cleanup after a folio error:
        1. Dismiss any <red-alert> ("no encontrado") overlay.
        2. Close the OC document modal if open.
        3. Close the Visualizar Factura modal if open.
        4. Close any remaining .redModal overlays.
        After this the search form should be accessible again.
        """
        try:
            await self._dismiss_red_alert(page)
        except Exception:
            pass

        for modal_id in ("content_modalDocOC", "content_modalVisualizar"):
            try:
                await self._close_modal(page, modal_id)
            except Exception:
                pass

        try:
            await self._close_all_modals(page)
        except Exception:
            pass

        await page.wait_for_timeout(400)

    @staticmethod
    async def _close_all_modals(page: Page):
        for _ in range(4):
            btns = page.locator(".redModal.ng-hide + .redModal .btn-cerrar25p:visible,"
                                " .redModal:not(.ng-hide) .btn-cerrar25p")
            if await btns.count() == 0:
                break
            await btns.last.click()
            await page.wait_for_timeout(400)
