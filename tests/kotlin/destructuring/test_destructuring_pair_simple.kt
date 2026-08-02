// vybe-test: kotlin/destructuring/test_destructuring_pair_simple
// origin: languages/kotlin/tests/kotlin/test_destructuring.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val pair = Pair(3, 4)
val (a, b) = pair
__check((a + b).toString(), "7") }
