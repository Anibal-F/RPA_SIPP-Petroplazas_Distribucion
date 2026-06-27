import asyncio
import flet as ft

def main(page: ft.Page):
    print(">>> main() called", flush=True)
    fp = ft.FilePicker(on_result=lambda e: print("RESULT:", e.files, flush=True))
    page.overlay.append(fp)
    page.add(ft.Text("Test"))
    page.update()

    async def trigger():
        await asyncio.sleep(2)
        print(">>> calling pick_files now", flush=True)
        fp.pick_files(dialog_title="TEST PICKER 0.28.2", allow_multiple=False)
        print(">>> pick_files call returned", flush=True)

    page.run_task(trigger)

ft.app(target=main)
