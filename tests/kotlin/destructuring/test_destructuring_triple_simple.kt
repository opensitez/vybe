// vybe-test: kotlin/destructuring/test_destructuring_triple_simple
// origin: languages/kotlin/tests/kotlin/test_destructuring.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val t = Triple(1, 2, 3)
val (a, b, c) = t
__check((a * b * c).toString(), "6") }
