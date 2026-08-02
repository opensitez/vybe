// vybe-test: kotlin/extension_functions/test_extension_property_getter
// origin: languages/kotlin/tests/kotlin/test_extension_functions.rs

class Point(val x: Int, val y: Int)

        val Point.sum: Int
            get() = x + y

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Point(2, 5).sum).toString(), "7")
        }
