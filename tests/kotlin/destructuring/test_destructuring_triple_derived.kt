// vybe-test: kotlin/destructuring/test_destructuring_triple_derived
// origin: languages/kotlin/tests/kotlin/test_destructuring.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val (x, y, z) = Triple(4, 5, 6)
__check((x + y + z).toString(), "15") }
