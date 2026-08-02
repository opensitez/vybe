// vybe-test: kotlin/destructuring/test_destructuring_from_local_function_return
// origin: languages/kotlin/tests/kotlin/test_destructuring.rs

fun coordinates(): Pair<Int, Int> {
            return Pair(9, 7)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val (x, y) = coordinates()
            __check((x + y).toString(), "16")
        }
