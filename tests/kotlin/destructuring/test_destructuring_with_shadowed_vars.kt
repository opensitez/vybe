// vybe-test: kotlin/destructuring/test_destructuring_with_shadowed_vars
// origin: languages/kotlin/tests/kotlin/test_destructuring.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pair1 = Pair(1, 2)
            val (a, b) = pair1
            val (c, d) = Pair(a + b, b + 1)
            __check((c).toString(), "3")
            __check((d).toString(), "3")
        }
