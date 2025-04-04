import asyncio
import flet as ft
import os
import sys
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))
from utils import config, get_app_size
from modules.pdf.process_pdf import process_pdfs
from modules.pdf.save_pdf import save_pdfs_from_folder, get_folder, connect_outlook

def main(page):
    assets_path = os.path.join(os.path.dirname(__file__), 'assets')
    win_width, win_height = get_app_size()
    config(page)
    
    folder_path = None
    error_snack_bar = None 
    state_image = ft.Image(src=os.path.join(assets_path, 'arrow-right.png'), width=70, height=70)

    def load_accounts():
        try:
            outlook = connect_outlook()
            accounts = []
            for account in outlook.Folders:
                accounts.append(ft.DropdownOption(text=account.Name))
            return accounts
        except Exception as e:
            print(f"Erro ao carregar contas: {e}")
            return []

    dropdown_options = load_accounts()

    async def process_folder(e):
        state_image.src = os.path.join(assets_path, 'loading-loop.gif')
        state_image.width = 50
        state_image.height = 50
        page.update()
        nonlocal folder_path
        folder_path = e.path
        if not folder_path:
            return
        folder_name_text.value = f"Pasta selecionada: {folder_path}"
        page.update()
        try:
            path = await process_pdfs(folder_path)
            show_success_dialog(f"Arquivo salvo em: {path}")
            state_image.src = os.path.join(assets_path, 'done-img.png')
        except Exception as error:
            show_error_snack_bar(str(error))
            state_image.src = os.path.join(assets_path, 'arrow-right.png')
            page.update()
        state_image.width = 50
        state_image.height = 50
        page.update()

    async def save_pdfs(e):
        save_pdfs_button.disabled = True  
        page.update()

        try:
            selected_account_name = dropdown.value
            print(f"Conta selecionada: {selected_account_name}")
            outlook = connect_outlook()
            selected_account = None
            for account in outlook.Folders:
                if account.Name == selected_account_name:
                    selected_account = account
                    break
            if not selected_account:
                raise Exception("Conta não encontrada.")
            folder = get_folder(outlook, "pdfs_medicoes")
            if not folder:
                raise Exception("Pasta 'pdfs_medicoes' não encontrada na conta selecionada.")
            state_image.src = os.path.join(assets_path, 'loading-loop.gif')
            state_image.width = 50
            state_image.height = 50
            page.update()
            await save_pdfs_from_folder()
            show_success_dialog("PDFs salvos com sucesso!")
            state_image.src = os.path.join(assets_path, 'done-img.png')
            page.update()
        except Exception as e:
            show_error_snack_bar(str(e))
            page.update()

        save_pdfs_button.disabled = False  
        page.update()

    save_pdfs_button = ft.Container(
        content=ft.ElevatedButton(
            "Salvar PDFs",
            on_click=save_pdfs,
            style=ft.ButtonStyle(
                bgcolor="#C0C0C0",
                color="black",
                shape=ft.RoundedRectangleBorder(radius=0),
                text_style=ft.TextStyle(size=18),
            ),
            width=200
        ),
        margin=ft.Margin(left=0, top=50, right=0, bottom=0)
    )

    def show_error_snack_bar(message):
        nonlocal error_snack_bar
        if error_snack_bar:
            error_snack_bar.open = False
            page.update()
        error_snack_bar = ft.SnackBar(content=ft.Text(message), bgcolor="red")
        page.overlay.append(error_snack_bar)
        error_snack_bar.open = True
        page.update()

    def show_success_dialog(message):
        snack_bar = ft.SnackBar(content=ft.Text(message), bgcolor="green")
        page.overlay.append(snack_bar)
        snack_bar.open = True
        page.update()

    async def dropdown_changed(e):
        selected_email = e.control.value
        print(f"Email selecionado: {selected_email}")
    
    folder_picker = ft.FilePicker(on_result=process_folder)

    open_file_button = ft.Container(
        content=ft.ElevatedButton(
            "Selecionar Pasta",
            on_click=lambda e: folder_picker.get_directory_path(),
            style=ft.ButtonStyle(
                bgcolor="#C0C0C0",
                color="black",
                shape=ft.RoundedRectangleBorder(radius=0),
                text_style=ft.TextStyle(size=18),
            ),
            width=200
        ),
        margin=ft.Margin(left=0, top=30, right=0, bottom=0)
    )

    folder_name_text = ft.Text(color='black', value="", size=16)
    
    top_bar = ft.Container(
        content=ft.Image(src=os.path.join(assets_path, 'bosch-bar.png'), fit=ft.ImageFit.COVER),
        width=win_width,
        height=8,
        padding=0,
        margin=0,
    )

    logo_image = ft.Container(
        content=ft.Image(src=os.path.join(assets_path, 'bosch_logo_basic.png'), width=150),
        margin=ft.Margin(bottom=60, top=0, left=0, right=0),
    )
    
    dropdown = ft.Dropdown(
        editable=False,
        options=dropdown_options,
        on_change=lambda e: asyncio.run(dropdown_changed(e)),
    )

    container = ft.Container(
        content=ft.Column(
            controls=[
                ft.Row(
                    controls=[
                        ft.Image(src=os.path.join(assets_path, 'pdf-img.svg'), width=100, height=100),
                        state_image,
                        ft.Image(src=os.path.join(assets_path, 'excel-img.svg'), width=100, height=100),
                    ],
                    alignment=ft.MainAxisAlignment.CENTER,
                ),
                save_pdfs_button,
                open_file_button,
                folder_picker,
                ft.Row(
                    controls=[folder_name_text],
                    alignment=ft.MainAxisAlignment.CENTER,
                ),
                ft.Row(
                    controls=[dropdown],
                    alignment=ft.MainAxisAlignment.CENTER,
                ),
            ],
            alignment=ft.MainAxisAlignment.CENTER,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            spacing=7
        ),
        bgcolor="#FFFFFF",
        border_radius=20,
    )

    top_bar_and_logo = ft.Column(
        controls=[top_bar, logo_image],
        alignment=ft.MainAxisAlignment.START,
        horizontal_alignment=ft.CrossAxisAlignment.START,
        spacing=0,
    )

    page.add(
        ft.Column(
            controls=[top_bar_and_logo, container],
            alignment=ft.MainAxisAlignment.START,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            height=win_height,
            width=win_width,
        )
    )
    
    page.update()

if __name__ == '__main__':
    ft.app(target=main)
