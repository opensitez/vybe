// vybe-test: kotlin/properties/test_property_initializer_order_with_dependency
// origin: languages/kotlin/tests/kotlin/test_properties.rs

class Grid {
            val width = 3
            val height = 4
            val area = width * height
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val grid = Grid()
            __check((grid.width).toString(), "3")
            __check((grid.height).toString(), "4")
            __check((grid.area).toString(), "12")
        }
