import sys
import os
import platform
import glob


def get_addins_path():
    """Devuelve la ruta de la carpeta de complementos de Excel según el sistema operativo."""
    if platform.system() == "Windows":
        return os.path.join(os.getenv("APPDATA"), "Microsoft", "AddIns")
    elif platform.system() == "Darwin":
        return os.path.join(
            os.path.expanduser("~"),
            "Library",
            "Group Containers",
            "UBF8T346G9.Office",
            "User Content",
            "Add-Ins",
        )
    else:
        return None


def find_addin(addin_name):
    """Busca un complemento XLAM en la carpeta de Add-ins de Excel."""
    addins_path = get_addins_path()

    if not addins_path:
        return None

    search_pattern = os.path.join(addins_path, f"*{addin_name}*.xlam")
    matches = glob.glob(search_pattern)

    if matches:
        # Si hay varios resultados, devolver el más reciente
        return max(matches, key=os.path.getctime)
    return None


def enable_addin(addin_path):
    """Activa un complemento XLAM en Excel."""
    try:
        if platform.system() == "Windows":
            import win32com.client
            import time

            try:
                # Iniciar Excel
                excel = win32com.client.DispatchEx("Excel.Application")
                excel.Visible = False  # Ejecutar en segundo plano
                excel.DisplayAlerts = False

                # Obtener solo el nombre del archivo
                addin_name = os.path.basename(addin_path)

                # _path = get_addins_path()

                # Obtener los complementos actuales
                addins = excel.AddIns
                is_active = False

                # Comprobar si el complemento ya está en la lista
                for addin in addins:
                    if (
                        addin.Name.lower() == addin_name.lower()
                    ):  # Compara la ruta completa
                        addin.Installed = True
                        is_active = True
                        print(f"✅ Complemento '{addin_name}' activado correctamente.")
                        break

                if not is_active:
                    print(
                        "❌ No fue posible activar el complemento de forma automática."
                    )

                # Esperar 5 segundos para ver Excel
                time.sleep(5)

                excel.Quit()

            except Exception as e:
                print(f"❌ Ocurrió un errorx: {e}")

        elif platform.system() == "Darwin":
            try:
                from appscript import app, mactypes

                excel = app("Microsoft Excel")
                addin_name = os.path.basename(addin_path)
                addin_path_abs = os.path.abspath(addin_path)
                file_ref = mactypes.File(addin_path_abs)

                if not os.path.exists(addin_path):
                    raise FileNotFoundError(f"Archivo no encontrado: {addin_path}")

                all_addins = excel.add_ins.name.get()
                print(all_addins)

                # Buscar por nombre exacto del complemento
                target_addin = next(
                    (
                        addin
                        for addin in all_addins
                        if getattr(addin.name, "get", lambda: "")() == addin_name
                    ),
                    None,
                )

                if not target_addin:
                    try:
                        # Método alternativo para instalar
                        excel.open(file_ref)
                        # Esperar y actualizar lista
                        import time

                        time.sleep(2)
                        all_addins = excel.add_ins.get()
                        target_addin = next(
                            (
                                addin
                                for addin in all_addins
                                if getattr(addin.name, "get", lambda: "")()
                                == addin_name
                            ),
                            None,
                        )
                    except Exception as install_error:
                        raise RuntimeError(f"Error instalando: {str(install_error)}")

                if target_addin:
                    target_addin.installed.set(True)
                    # Verificar activación
                    if target_addin.installed.get():
                        print(f"ÉXITO: {addin_name} activado")
                    else:
                        print(f"ADVERTENCIA: {addin_name} no se pudo activar")
                else:
                    print("ERROR: Complemento no encontrado tras instalación")

            except Exception as e:
                print(f"ERROR: {str(e)}")

    except Exception as e:
        return f"ERROR: {str(e)}"


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("ERROR: Debes proporcionar el nombre del complemento.")
        sys.exit(1)

    addin_name = sys.argv[1]
    # addin_name = "plantilla"
    addin_path = find_addin(addin_name)

    if addin_path:
        print(enable_addin(addin_path))
    else:
        print(f"ERROR: No se encontró ningún complemento con el nombre '{addin_name}'")


# En Windows
# pyinstaller --onefile --noconsole --name AddInManager.exe index.py

# En macOS
# pyinstaller --onefile --noconsole --name AddInManager excel_addin.py

