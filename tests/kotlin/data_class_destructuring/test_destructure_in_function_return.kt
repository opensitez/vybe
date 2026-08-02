// vybe-test: kotlin/data_class_destructuring/test_destructure_in_function_return
// origin: languages/kotlin/tests/kotlin/test_data_class_destructuring.rs

data class Point(val x: Int, val y: Int)

        fun origin(): Point = Point(0, 0)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val (x, y) = origin()
            __check((x).toString(), "0")
            __check((y).toString(), "0")
        }
