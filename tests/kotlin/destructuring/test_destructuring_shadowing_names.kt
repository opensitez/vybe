// vybe-test: kotlin/destructuring/test_destructuring_shadowing_names
// origin: languages/kotlin/tests/kotlin/test_destructuring.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val (a, b) = Pair(7, 8)
val (a2, b2) = Pair(a + 1, b + 1)
__check((a2).toString(), "8")
__check((b2).toString(), "9") }
