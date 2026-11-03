import flet as ft
import json
import os
from datetime import datetime
import uuid

try:
    import openpyxl
    from openpyxl import Workbook
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

DB_FILE = 'os_offline.json'
CONFIG_FILE = 'config.json'

def carregar_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {"equipe": "Minha Equipe", "caminho_planilha": ""}

def salvar_config(equipe, caminho):
    config = {"equipe": equipe.strip(), "caminho_planilha": caminho.strip()}
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=4)

def carregar_os():
    if os.path.exists(DB_FILE):
        with open(DB_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return []

def salvar_os(dados):
    with open(DB_FILE, 'w', encoding='utf-8') as f:
        json.dump(dados, f, indent=4)

def main(page: ft.Page):
    page.title = "TROCA DE EQUIPAMENTOS"
    page.theme_mode = ft.ThemeMode.LIGHT
    page.window_width = 400
    page.window_height = 800
    page.bgcolor = "#F3F4F6"
    
    app_state = {
        "os_atual": {},
        "config": carregar_config()
    }

    def route_change(route):
        page.views.clear()
        
        # --- TELA: MENU PRINCIPAL ---
        if page.route == "/":
            oss = carregar_os()
            pendentes = sum(1 for o in oss if not o.get("sincronizada", False))
            
            alerta_text = f"🔔 {pendentes} OS(s) PENDENTES" if pendentes > 0 else "✅ TUDO SINCRONIZADO"
            alerta_color = ft.colors.RED_700 if pendentes > 0 else ft.colors.GREEN_700
            alerta_bg = ft.colors.RED_100 if pendentes > 0 else ft.colors.GREEN_100

            page.views.append(
                ft.View(
                    "/",
                    [
                        ft.Container(
                            content=ft.Column([
                                ft.Text("📱 TROCA DE EQUIPAMENTOS", size=20, weight=ft.FontWeight.BOLD, color=ft.colors.WHITE),
                                ft.Text("Painel de Controle Mobile", size=12, color=ft.colors.BLUE_100)
                            ]),
                            bgcolor=ft.colors.BLUE_700, padding=20, width=float('inf')
                        ),
                        ft.Container(
                            content=ft.Text(alerta_text, weight=ft.FontWeight.BOLD, color=alerta_color),
                            bgcolor=alerta_bg, padding=15, border_radius=8, alignment=ft.alignment.center, margin=ft.margin.symmetric(horizontal=20, vertical=10)
                        ),
                        ft.Column([
                            ft.ElevatedButton("⚡ NOVA OS", on_click=lambda _: iniciar_nova_os(), width=float('inf'), height=50, style=ft.ButtonStyle(bgcolor=ft.colors.BLUE_700, color=ft.colors.WHITE)),
                            ft.ElevatedButton("📊 DASHBOARD", on_click=lambda _: page.go("/dashboard"), width=float('inf'), height=50, style=ft.ButtonStyle(bgcolor=ft.colors.TEAL_700, color=ft.colors.WHITE)),
                            ft.ElevatedButton("📋 HISTÓRICO", on_click=lambda _: page.go("/historico"), width=float('inf'), height=50, style=ft.ButtonStyle(bgcolor=ft.colors.WHITE, color=ft.colors.BLACK87)),
                            ft.ElevatedButton("⚙️ CONFIGURAÇÕES", on_click=lambda _: page.go("/config"), width=float('inf'), height=50, style=ft.ButtonStyle(bgcolor=ft.colors.GREY_200, color=ft.colors.BLACK87)),
                            ft.ElevatedButton("🔄 ATUALIZAR EXCEL", on_click=lambda _: exportar_excel(), width=float('inf'), height=50, style=ft.ButtonStyle(bgcolor=ft.colors.GREEN_700, color=ft.colors.WHITE)),
                        ], spacing=10, width=float('inf'))
                    ],
                    scroll=ft.ScrollMode.AUTO,
                    padding=0
                )
            )

        # --- TELA: NOVA OS (DADOS BÁSICOS) ---
        elif page.route == "/nova_os":
            inst_field = ft.TextField(label="Instalação do Cliente", value=app_state["os_atual"].get("instalacao", ""))
            nota_field = ft.TextField(label="Número da Nota", value=app_state["os_atual"].get("nota", ""))
            tipo_dd = ft.Dropdown(
                label="Tipo de Serviço",
                options=[ft.dropdown.Option(x) for x in ["Migração ACL", "Normalização", "Modernização", "Cativo", "Autoprodutor", "Outros"]],
                value=app_state["os_atual"].get("tipo_servico")
            )
            retaguarda_check = ft.Switch(label="Possui medidor retaguarda?", value=app_state["os_atual"].get("retaguarda") == "Sim")

            def avancar_os(e):
                if not inst_field.value or not nota_field.value or not tipo_dd.value:
                    page.snack_bar = ft.SnackBar(ft.Text("Preencha Instalação, Nota e Serviço."))
                    page.snack_bar.open = True
                    page.update()
                    return
                app_state["os_atual"]["instalacao"] = inst_field.value
                app_state["os_atual"]["nota"] = nota_field.value
                app_state["os_atual"]["tipo_servico"] = tipo_dd.value
                app_state["os_atual"]["retaguarda"] = "Sim" if retaguarda_check.value else "Não"
                page.go("/equipamentos")

            page.views.append(
                ft.View(
                    "/nova_os",
                    [
                        ft.AppBar(title=ft.Text("NOVA OS"), bgcolor=ft.colors.BLUE_700, color=ft.colors.WHITE),
                        ft.Container(
                            content=ft.Column([
                                ft.Text(f"Código: {app_state['os_atual'].get('codigo_os', '')}", weight=ft.FontWeight.BOLD, color=ft.colors.BLUE_700),
                                inst_field, nota_field, tipo_dd, retaguarda_check,
                                ft.ElevatedButton("AVANÇAR ➔", on_click=avancar_os, width=float('inf'), height=50, style=ft.ButtonStyle(bgcolor=ft.colors.BLUE_700, color=ft.colors.WHITE))
                            ], spacing=20),
                            padding=20
                        )
                    ], scroll=ft.ScrollMode.AUTO
                )
            )

        # --- TELA: EQUIPAMENTOS E LEITURA ---
        elif page.route == "/equipamentos":
            tem_retaguarda = app_state["os_atual"].get("retaguarda") == "Sim"
            e_cativo = app_state["os_atual"].get("tipo_servico") == "Cativo"
            
            opcoes_todas = [
                "MEDIDOR RETIRADO", "MEDIDOR NOVO", "REMOTA RETIRADA", "REMOTA NOVA",
                "SSN EDP RETIRADO", "SSN EDP NOVO", "SSN CCEE RETIRADO", "SSN CCEE NOVO",
                "MEDIDOR RETAGUARDA RETIRADO", "MEDIDOR RETAGUARDA NOVO",
                "REMOTA RETAGUARDA RETIRADA", "REMOTA RETAGUARDA NOVA",
                "SSN EDP RETAGUARDA RETIRADO", "SSN EDP RETAGUARDA NOVO",
                "SSN CCEE RETAGUARDA RETIRADO", "SSN CCEE RETAGUARDA NOVO"
            ]
            
            campos_ativos = [i for i in opcoes_todas if not (not tem_retaguarda and "RETAGUARDA" in i) and not (e_cativo and "CCEE" in i)]
            
            fields_dict = {}
            col_controls = []
            
            # Lógica para pular o foco com o Enter do leitor
            def criar_on_submit(index):
                def handler(e):
                    if index + 1 < len(field_objects):
                        field_objects[index + 1].focus()
                    else:
                        obs_field.focus()
                return handler

            field_objects = []
            for equip in campos_ativos:
                tf = ft.TextField(
                    label=equip, 
                    value=app_state["os_atual"].get("equipamentos", {}).get(equip, ""),
                    text_size=14
                )
                fields_dict[equip] = tf
                field_objects.append(tf)
                col_controls.append(tf)

            for i, tf in enumerate(field_objects):
                tf.on_submit = criar_on_submit(i)

            constante_field = ft.TextField(label="Constante", value=app_state["os_atual"].get("constante", ""), visible=False)
            constante_switch = ft.Switch(label="Preencher constante?", value=bool(app_state["os_atual"].get("constante", "")))
            
            def toggle_constante(e):
                constante_field.visible = constante_switch.value
                page.update()
                
            constante_switch.on_change = toggle_constante
            if constante_switch.value: constante_field.visible = True

            if "MEDIDOR NOVO" in campos_ativos:
                col_controls.append(constante_switch)
                col_controls.append(constante_field)

            obs_field = ft.TextField(label="Observações", multiline=True, min_lines=2, value=app_state["os_atual"].get("observacoes", ""))
            col_controls.append(obs_field)

            def finalizar_os(e):
                equipamentos = {k: v.value.strip() for k, v in fields_dict.items() if v.value.strip()}
                if not equipamentos:
                    page.snack_bar = ft.SnackBar(ft.Text("Aviso: Nenhum equipamento preenchido."))
                    page.snack_bar.open = True
                    page.update()
                    
                app_state["os_atual"]["equipamentos"] = equipamentos
                app_state["os_atual"]["constante"] = constante_field.value.strip() if constante_switch.value else ""
                app_state["os_atual"]["observacoes"] = obs_field.value.strip()
                
                oss = carregar_os()
                atualizado = False
                for i, o in enumerate(oss):
                    if o["codigo_os"] == app_state["os_atual"]["codigo_os"]:
                        oss[i] = app_state["os_atual"]
                        atualizado = True
                        break
                if not atualizado:
                    oss.append(app_state["os_atual"])
                    
                salvar_os(oss)
                page.snack_bar = ft.SnackBar(ft.Text("OS Salva com sucesso!"))
                page.snack_bar.open = True
                page.go("/")

            col_controls.append(ft.ElevatedButton("FINALIZAR E SALVAR OS", on_click=finalizar_os, width=float('inf'), height=50, style=ft.ButtonStyle(bgcolor=ft.colors.GREEN_700, color=ft.colors.WHITE)))

            page.views.append(
                ft.View(
                    "/equipamentos",
                    [
                        ft.AppBar(title=ft.Text("EQUIPAMENTOS"), bgcolor=ft.colors.BLUE_700, color=ft.colors.WHITE),
                        ft.Container(content=ft.Column(col_controls, spacing=15), padding=20)
                    ], scroll=ft.ScrollMode.AUTO
                )
            )

        # --- TELA: DASHBOARD ---
        elif page.route == "/dashboard":
            oss = carregar_os()
            notas_dia, notas_semana, notas_mes = 0, 0, 0
            agora = datetime.now()
            hoje_str = agora.strftime('%d/%m/%Y')
            mes_atual = agora.strftime('%m/%Y')
            semana_atual = agora.isocalendar()[1]
            
            for o in oss:
                dt_str = o.get("data_criacao", "")
                try:
                    dt_obj = datetime.strptime(dt_str[:10], '%d/%m/%Y')
                    if dt_str.startswith(hoje_str): notas_dia += 1
                    if dt_str[3:10] == mes_atual: notas_mes += 1
                    if dt_obj.isocalendar()[1] == semana_atual and dt_obj.year == agora.year: notas_semana += 1
                except:
                    pass

            def create_card(titulo, valor, cor):
                return ft.Container(
                    content=ft.Column([
                        ft.Text(titulo, weight=ft.FontWeight.BOLD, color=ft.colors.GREY_600),
                        ft.Text(str(valor), size=30, weight=ft.FontWeight.BOLD, color=cor)
                    ], alignment=ft.MainAxisAlignment.CENTER, horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                    bgcolor=ft.colors.WHITE, padding=20, border_radius=8, border=ft.border.all(1, ft.colors.GREY_300),
                    width=float('inf')
                )

            page.views.append(
                ft.View(
                    "/dashboard",
                    [
                        ft.AppBar(title=ft.Text("DASHBOARD"), bgcolor=ft.colors.TEAL_700, color=ft.colors.WHITE),
                        ft.Container(
                            content=ft.Column([
                                create_card("NOTAS DO DIA", notas_dia, ft.colors.BLUE_700),
                                create_card("NOTAS DA SEMANA", notas_semana, ft.colors.TEAL_700),
                                create_card("NOTAS DO MÊS", notas_mes, ft.colors.GREEN_700),
                            ], spacing=15), padding=20
                        )
                    ], scroll=ft.ScrollMode.AUTO
                )
            )

        # --- TELA: HISTÓRICO ---
        elif page.route == "/historico":
            oss = reversed(carregar_os())
            lista_cards = []
            for o in oss:
                sinc = o.get("sincronizada", False)
                bg_color = ft.colors.GREEN_50 if sinc else ft.colors.RED_50
                icon_color = ft.colors.GREEN_700 if sinc else ft.colors.RED_700
                status = "SINCRONIZADA" if sinc else "PENDENTE"
                
                lista_cards.append(
                    ft.Container(
                        content=ft.Column([
                            ft.Text(f"Inst: {o.get('instalacao')}", weight=ft.FontWeight.BOLD),
                            ft.Text(f"Data: {o.get('data_criacao')} - {status}", color=icon_color, size=12)
                        ]),
                        bgcolor=bg_color, padding=15, border_radius=8, margin=ft.margin.only(bottom=10)
                    )
                )

            page.views.append(
                ft.View(
                    "/historico",
                    [
                        ft.AppBar(title=ft.Text("HISTÓRICO"), bgcolor=ft.colors.BLUE_700, color=ft.colors.WHITE),
                        ft.Container(content=ft.Column(lista_cards), padding=20)
                    ], scroll=ft.ScrollMode.AUTO
                )
            )

        # --- TELA: CONFIGURAÇÕES ---
        elif page.route == "/config":
            equipe_field = ft.TextField(label="Nome da Equipe", value=app_state["config"].get("equipe", ""))
            pasta_field = ft.TextField(label="Caminho/Pasta do Excel Mobile", value=app_state["config"].get("caminho_planilha", ""))
            
            def salvar_cfg(e):
                salvar_config(equipe_field.value, pasta_field.value)
                app_state["config"] = carregar_config()
                page.snack_bar = ft.SnackBar(ft.Text("Configurações salvas!"))
                page.snack_bar.open = True
                page.go("/")

            page.views.append(
                ft.View(
                    "/config",
                    [
                        ft.AppBar(title=ft.Text("CONFIGURAÇÕES"), bgcolor=ft.colors.GREY_800, color=ft.colors.WHITE),
                        ft.Container(
                            content=ft.Column([
                                equipe_field, pasta_field,
                                ft.Text("*No Android, pastas públicas usam caminhos como /storage/emulated/0/Download", size=11, color=ft.colors.GREY_500),
                                ft.ElevatedButton("SALVAR", on_click=salvar_cfg, width=float('inf'), height=50, style=ft.ButtonStyle(bgcolor=ft.colors.BLUE_700, color=ft.colors.WHITE))
                            ], spacing=20), padding=20
                        )
                    ], scroll=ft.ScrollMode.AUTO
                )
            )

        page.update()

    def iniciar_nova_os():
        app_state["os_atual"] = {
            "codigo_os": f"OS-{datetime.now().strftime('%Y%m%d%H%M%S')}-{str(uuid.uuid4())[:4].upper()}",
            "data_criacao": datetime.now().strftime('%d/%m/%Y %H:%M:%S'),
            "sincronizada": False,
            "equipamentos": {},
            "constante": "",
            "observacoes": ""
        }
        page.go("/nova_os")

    def exportar_excel():
        if not HAS_OPENPYXL:
            page.snack_bar = ft.SnackBar(ft.Text("Erro: Biblioteca openpyxl não instalada."))
            page.snack_bar.open = True
            page.update()
            return
            
        pasta = app_state["config"].get("caminho_planilha", "").strip()
        eq = app_state["config"].get("equipe", "Equipe").strip()
        
        if not pasta:
            page.snack_bar = ft.SnackBar(ft.Text("Configure a pasta de destino nas configurações."))
            page.snack_bar.open = True
            page.update()
            return
            
        oss = carregar_os()
        pendentes = [o for o in oss if not o.get("sincronizada", False)]
        if not pendentes:
            page.snack_bar = ft.SnackBar(ft.Text("Nenhuma OS pendente de sincronização."))
            page.snack_bar.open = True
            page.update()
            return
            
        arquivo = os.path.join(pasta, f"{eq}.xlsx")
        cabecalho = ["CÓDIGO OS", "DATA", "INSTALAÇÃO", "NOTA", "TIPO", "CONSTANTE", "MED RET", "MED NOVO", "REM RET", "REM NOVA", "OBS"]

        try:
            wb = openpyxl.load_workbook(arquivo) if os.path.exists(arquivo) else Workbook()
            ws = wb.active
            if not os.path.exists(arquivo): ws.append(cabecalho)
            
            for d in pendentes:
                eqp = d.get("equipamentos", {})
                ws.append([
                    d.get("codigo_os"), d.get("data_criacao"), d.get("instalacao"), d.get("nota"), d.get("tipo_servico"), d.get("constante"),
                    eqp.get("MEDIDOR RETIRADO", ""), eqp.get("MEDIDOR NOVO", ""), eqp.get("REMOTA RETIRADA", ""), eqp.get("REMOTA NOVA", ""),
                    d.get("observacoes", "")
                ])
                d["sincronizada"] = True
                
            wb.save(arquivo)
            salvar_os(oss)
            page.snack_bar = ft.SnackBar(ft.Text(f"{len(pendentes)} OS(s) exportadas com sucesso!"))
            page.snack_bar.open = True
            page.go("/")
        except Exception as e:
            page.snack_bar = ft.SnackBar(ft.Text(f"Erro ao salvar: {str(e)}"))
            page.snack_bar.open = True
            page.update()

    def view_pop(view):
        page.views.pop()
        top_view = page.views[-1]
        page.go(top_view.route)

    page.on_route_change = route_change
    page.on_view_pop = view_pop
    page.go(page.route)

ft.app(target=main)
