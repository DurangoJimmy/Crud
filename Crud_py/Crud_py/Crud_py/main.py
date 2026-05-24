"""
main.py — Punto de entrada con menú interactivo para gestionar productos.
"""
from productos import ProductoCRUD


def mostrar_producto(p: dict):
    print(
        f"  ID: {p['id']} | Nombre: {p['nombre']} | "
        f"Precio: ${p['precio']:.2f} | Cantidad: {p['cantidad']}\n"
        f"  Descripción: {p['descripcion']}"
    )


def menu():
    crud = ProductoCRUD(archivo="productos.json")

    opciones = {
        "1": "Crear producto",
        "2": "Ver producto por ID",
        "3": "Listar todos los productos",
        "4": "Actualizar producto",
        "5": "Eliminar producto",
        "0": "Salir",
    }

    while True:
        print("\n========= Gestión de Productos =========")
        for clave, desc in opciones.items():
            print(f"  {clave}. {desc}")
        print("=========================================")
        opcion = input("Seleccione una opción: ").strip()

        if opcion == "1":
            try:
                id_prod = int(input("  ID: "))
                nombre = input("  Nombre: ").strip()
                descripcion = input("  Descripción: ").strip()
                precio = float(input("  Precio: "))
                cantidad = int(input("  Cantidad: "))
                p = crud.crear(id_prod, nombre, descripcion, precio, cantidad)
                print("[OK] Producto creado:")
                mostrar_producto(p)
            except (ValueError, KeyError) as e:
                print(f"[ERROR] {e}")

        elif opcion == "2":
            try:
                id_prod = int(input("  ID: "))
                p = crud.leer(id_prod)
                mostrar_producto(p)
            except (ValueError, KeyError) as e:
                print(f"[ERROR] {e}")

        elif opcion == "3":
            productos = crud.leer_todos()
            if not productos:
                print("No hay productos registrados.")
            else:
                for p in productos:
                    mostrar_producto(p)
                    print()

        elif opcion == "4":
            try:
                id_prod = int(input("  ID del producto a actualizar: "))
                print("  (Deje en blanco los campos que no desee cambiar)")
                campos = {}
                nombre = input("  Nuevo nombre: ").strip()
                if nombre:
                    campos["nombre"] = nombre
                descripcion = input("  Nueva descripción: ").strip()
                if descripcion:
                    campos["descripcion"] = descripcion
                precio = input("  Nuevo precio: ").strip()
                if precio:
                    campos["precio"] = float(precio)
                cantidad = input("  Nueva cantidad: ").strip()
                if cantidad:
                    campos["cantidad"] = int(cantidad)
                if not campos:
                    print("[INFO] No se modificó nada.")
                else:
                    p = crud.actualizar(id_prod, **campos)
                    print("[OK] Producto actualizado:")
                    mostrar_producto(p)
            except (ValueError, KeyError) as e:
                print(f"[ERROR] {e}")

        elif opcion == "5":
            try:
                id_prod = int(input("  ID del producto a eliminar: "))
                p = crud.eliminar(id_prod)
                print(f"[OK] Producto eliminado: {p['nombre']} (ID={p['id']})")
            except (ValueError, KeyError) as e:
                print(f"[ERROR] {e}")

        elif opcion == "0":
            print("Hasta luego.")
            break

        else:
            print("[INFO] Opción no válida, intente de nuevo.")


if __name__ == "__main__":
    menu()
