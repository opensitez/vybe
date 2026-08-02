// vybe-test: kotlin/data_classes/test_data_class_plus_operator_style_via_copy
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Coord(val x: Int, val y: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val origin = Coord(0, 0)
            fun move(point: Coord, dx: Int, dy: Int): Coord {
                return point.copy(x = point.x + dx, y = point.y + dy)
            }
            val moved = move(origin, 3, 4)
            __check((moved.x).toString(), "3")
            __check((moved.y).toString(), "4")
            __check((origin.x).toString(), "0")
        }
