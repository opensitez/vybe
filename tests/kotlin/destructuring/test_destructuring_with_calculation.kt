// vybe-test: kotlin/destructuring/test_destructuring_with_calculation
// origin: languages/kotlin/tests/kotlin/test_destructuring.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val (a, b) = Pair(8, 3)
            __check((a - b).toString(), "5")
            __check((a + b).toString(), "11")
            __check((a * b).toString(), "24")
        }
