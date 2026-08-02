// vybe-test: kotlin/destructuring/test_destructuring_ignores_unused_positions
// origin: languages/kotlin/tests/kotlin/test_destructuring.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val (first, _, third) = Triple("one", "skip", "three")
            __check((first).toString(), "one")
            __check((third).toString(), "three")
        }
