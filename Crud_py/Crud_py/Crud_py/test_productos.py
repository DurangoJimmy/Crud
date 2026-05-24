"""
test_productos.py — Pruebas unitarias para las operaciones CRUD de productos.
Cada operación tiene al menos un caso exitoso y uno de error.
"""
import unittest
from productos import ProductoCRUD


# ═══════════════════════════════════════════════════════════════════════════ #
# CREAR                                                                       #
# ═══════════════════════════════════════════════════════════════════════════ #
class TestCrear(unittest.TestCase):

    def setUp(self):
        self.crud = ProductoCRUD()

    # -- casos exitosos --
    def test_crear_producto_retorna_datos_correctos(self):
        p = self.crud.crear(1, "Laptop", "Computadora portátil", 2500.0, 10)
        self.assertEqual(p["id"], 1)
        self.assertEqual(p["nombre"], "Laptop")
        self.assertEqual(p["descripcion"], "Computadora portátil")
        self.assertEqual(p["precio"], 2500.0)
        self.assertEqual(p["cantidad"], 10)

    def test_crear_producto_queda_almacenado(self):
        self.crud.crear(2, "Mouse", "Periférico de entrada", 50.0, 25)
        p = self.crud.leer(2)
        self.assertEqual(p["nombre"], "Mouse")

    # -- casos de error --
    def test_crear_id_duplicado_lanza_ValueError(self):
        self.crud.crear(1, "Laptop", "Desc", 2500.0, 10)
        with self.assertRaises(ValueError):
            self.crud.crear(1, "Mouse", "Desc2", 50.0, 5)

    def test_crear_precio_negativo_lanza_ValueError(self):
        with self.assertRaises(ValueError):
            self.crud.crear(3, "Teclado", "Periférico", -100.0, 5)

    def test_crear_cantidad_negativa_lanza_ValueError(self):
        with self.assertRaises(ValueError):
            self.crud.crear(4, "Monitor", "Pantalla 24\"", 500.0, -1)


# ═══════════════════════════════════════════════════════════════════════════ #
# LEER                                                                        #
# ═══════════════════════════════════════════════════════════════════════════ #
class TestLeer(unittest.TestCase):

    def setUp(self):
        self.crud = ProductoCRUD()
        self.crud.crear(1, "Laptop", "Computadora portátil", 2500.0, 10)
        self.crud.crear(2, "Mouse", "Periférico de entrada", 50.0, 25)

    # -- casos exitosos --
    def test_leer_producto_existente(self):
        p = self.crud.leer(1)
        self.assertEqual(p["nombre"], "Laptop")

    def test_leer_todos_devuelve_lista_completa(self):
        todos = self.crud.leer_todos()
        self.assertEqual(len(todos), 2)

    def test_leer_todos_vacio_devuelve_lista_vacia(self):
        crud_vacio = ProductoCRUD()
        self.assertEqual(crud_vacio.leer_todos(), [])

    # -- casos de error --
    def test_leer_id_inexistente_lanza_KeyError(self):
        with self.assertRaises(KeyError):
            self.crud.leer(999)

    def test_leer_devuelve_copia_independiente(self):
        """Modificar el dict retornado no debe alterar el almacenamiento."""
        p = self.crud.leer(1)
        p["nombre"] = "Modificado externamente"
        self.assertEqual(self.crud.leer(1)["nombre"], "Laptop")


# ═══════════════════════════════════════════════════════════════════════════ #
# ACTUALIZAR                                                                  #
# ═══════════════════════════════════════════════════════════════════════════ #
class TestActualizar(unittest.TestCase):

    def setUp(self):
        self.crud = ProductoCRUD()
        self.crud.crear(1, "Laptop", "Computadora portátil", 2500.0, 10)

    # -- casos exitosos --
    def test_actualizar_precio_y_cantidad(self):
        p = self.crud.actualizar(1, precio=2000.0, cantidad=8)
        self.assertEqual(p["precio"], 2000.0)
        self.assertEqual(p["cantidad"], 8)

    def test_actualizar_nombre(self):
        p = self.crud.actualizar(1, nombre="Laptop Pro")
        self.assertEqual(p["nombre"], "Laptop Pro")

    def test_actualizar_descripcion(self):
        p = self.crud.actualizar(1, descripcion="Nueva descripción")
        self.assertEqual(p["descripcion"], "Nueva descripción")

    # -- casos de error --
    def test_actualizar_producto_inexistente_lanza_KeyError(self):
        with self.assertRaises(KeyError):
            self.crud.actualizar(999, nombre="Fantasma")

    def test_actualizar_precio_negativo_lanza_ValueError(self):
        with self.assertRaises(ValueError):
            self.crud.actualizar(1, precio=-50.0)

    def test_actualizar_cantidad_negativa_lanza_ValueError(self):
        with self.assertRaises(ValueError):
            self.crud.actualizar(1, cantidad=-5)

    def test_actualizar_campo_invalido_lanza_ValueError(self):
        with self.assertRaises(ValueError):
            self.crud.actualizar(1, color="rojo")


# ═══════════════════════════════════════════════════════════════════════════ #
# ELIMINAR                                                                    #
# ═══════════════════════════════════════════════════════════════════════════ #
class TestEliminar(unittest.TestCase):

    def setUp(self):
        self.crud = ProductoCRUD()
        self.crud.crear(1, "Laptop", "Computadora portátil", 2500.0, 10)
        self.crud.crear(2, "Mouse", "Periférico de entrada", 50.0, 25)

    # -- casos exitosos --
    def test_eliminar_producto_existente(self):
        eliminado = self.crud.eliminar(1)
        self.assertEqual(eliminado["id"], 1)

    def test_eliminar_reduce_conteo(self):
        self.crud.eliminar(1)
        self.assertEqual(len(self.crud.leer_todos()), 1)

    def test_eliminar_producto_ya_no_accesible(self):
        self.crud.eliminar(1)
        with self.assertRaises(KeyError):
            self.crud.leer(1)

    # -- casos de error --
    def test_eliminar_producto_inexistente_lanza_KeyError(self):
        with self.assertRaises(KeyError):
            self.crud.eliminar(999)

    def test_eliminar_mismo_producto_dos_veces_lanza_KeyError(self):
        self.crud.eliminar(2)
        with self.assertRaises(KeyError):
            self.crud.eliminar(2)


if __name__ == "__main__":
    unittest.main(verbosity=2)
