"""
productos.py — Lógica CRUD para el registro de productos.
Almacenamiento: diccionario en memoria con persistencia opcional en JSON.
"""
import json
import os


class ProductoCRUD:
    """Gestiona las operaciones CRUD sobre un registro de productos."""

    def __init__(self, archivo: str = None):
        """
        Args:
            archivo: Ruta del archivo JSON para persistencia. Si es None,
                     los datos solo se guardan en memoria.
        """
        self._productos: dict = {}
        self._archivo = archivo
        if archivo and os.path.exists(archivo):
            self._cargar()

    # ------------------------------------------------------------------ #
    # CREATE                                                               #
    # ------------------------------------------------------------------ #
    def crear(
        self,
        id: int,
        nombre: str,
        descripcion: str,
        precio: float,
        cantidad: int,
    ) -> dict:
        """Crea un nuevo producto.

        Raises:
            ValueError: Si el id ya existe, el precio o la cantidad son negativos.
        """
        if id in self._productos:
            raise ValueError(f"Ya existe un producto con id={id}")
        if precio < 0:
            raise ValueError("El precio no puede ser negativo")
        if cantidad < 0:
            raise ValueError("La cantidad no puede ser negativa")

        producto = {
            "id": id,
            "nombre": nombre,
            "descripcion": descripcion,
            "precio": precio,
            "cantidad": cantidad,
        }
        self._productos[id] = producto
        self._guardar()
        return dict(producto)

    # ------------------------------------------------------------------ #
    # READ                                                                 #
    # ------------------------------------------------------------------ #
    def leer(self, id: int) -> dict:
        """Devuelve el producto con el id indicado.

        Raises:
            KeyError: Si no existe ningún producto con ese id.
        """
        if id not in self._productos:
            raise KeyError(f"No existe producto con id={id}")
        return dict(self._productos[id])

    def leer_todos(self) -> list:
        """Devuelve una lista con todos los productos."""
        return [dict(p) for p in self._productos.values()]

    # ------------------------------------------------------------------ #
    # UPDATE                                                               #
    # ------------------------------------------------------------------ #
    def actualizar(self, id: int, **campos) -> dict:
        """Actualiza uno o más campos de un producto existente.

        Raises:
            KeyError: Si no existe ningún producto con ese id.
            ValueError: Si se pasa un campo no permitido o un valor inválido.
        """
        if id not in self._productos:
            raise KeyError(f"No existe producto con id={id}")

        campos_validos = {"nombre", "descripcion", "precio", "cantidad"}
        invalidos = set(campos) - campos_validos
        if invalidos:
            raise ValueError(f"Campos no válidos: {invalidos}")

        if "precio" in campos and campos["precio"] < 0:
            raise ValueError("El precio no puede ser negativo")
        if "cantidad" in campos and campos["cantidad"] < 0:
            raise ValueError("La cantidad no puede ser negativa")

        self._productos[id].update(campos)
        self._guardar()
        return dict(self._productos[id])

    # ------------------------------------------------------------------ #
    # DELETE                                                               #
    # ------------------------------------------------------------------ #
    def eliminar(self, id: int) -> dict:
        """Elimina el producto con el id indicado y lo devuelve.

        Raises:
            KeyError: Si no existe ningún producto con ese id.
        """
        if id not in self._productos:
            raise KeyError(f"No existe producto con id={id}")
        eliminado = self._productos.pop(id)
        self._guardar()
        return dict(eliminado)

    # ------------------------------------------------------------------ #
    # Persistencia JSON                                                    #
    # ------------------------------------------------------------------ #
    def _guardar(self):
        if self._archivo:
            with open(self._archivo, "w", encoding="utf-8") as f:
                json.dump(list(self._productos.values()), f, ensure_ascii=False, indent=2)

    def _cargar(self):
        with open(self._archivo, "r", encoding="utf-8") as f:
            datos = json.load(f)
        self._productos = {p["id"]: p for p in datos}
