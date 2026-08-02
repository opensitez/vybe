// vybe-test: kotlin/destructuring/test_destructuring_multiple_calls
// origin: languages/kotlin/tests/kotlin/test_destructuring.rs

fun combine(first: Pair<Int, Int>): Int { val (a, b) = first
return a + b }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((combine(Pair(2, 8))).toString(), "10")
__check((combine(Pair(1, 2))).toString(), "3") }
