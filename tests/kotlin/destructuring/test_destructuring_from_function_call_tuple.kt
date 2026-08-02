// vybe-test: kotlin/destructuring/test_destructuring_from_function_call_tuple
// origin: languages/kotlin/tests/kotlin/test_destructuring.rs

fun coordinates(): Pair<String, Pair<Int, Int>> = Pair("pt", Pair(7, 8))

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val (name, point) = coordinates()
            val (x, y) = point
            __check((name).toString(), "pt")
            __check((x).toString(), "7")
            __check((y).toString(), "8")
        }
