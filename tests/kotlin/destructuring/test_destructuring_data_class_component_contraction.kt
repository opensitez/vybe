// vybe-test: kotlin/destructuring/test_destructuring_data_class_component_contraction
// origin: languages/kotlin/tests/kotlin/test_destructuring.rs

data class Point(val x: Int, val y: Int, val z: Int)

        fun origin(): Point = Point(2, 4, 6)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val (x, y, z) = origin()
            __check((x + y).toString(), "6")
            __check((z - y).toString(), "2")
        }
