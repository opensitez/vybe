// vybe-test: kotlin/destructuring/test_destructuring_chained_expression
// origin: languages/kotlin/tests/kotlin/test_destructuring.rs

fun makePair(v: Int): Pair<Int, Int> { return Pair(v, v + 1) }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val p = makePair(10)
val (x, y) = p
__check((y - x).toString(), "1") }
