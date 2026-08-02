// vybe-test: kotlin/destructuring/test_destructuring_nested_return
// origin: languages/kotlin/tests/kotlin/test_destructuring.rs

fun bundle(): Pair<Pair<Int, Int>, Pair<Int, Int>> = Pair(Pair(1, 2), Pair(3, 4))
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val (first, second) = bundle()
val (a, b) = first
val (c, d) = second
__check((a + b + c + d).toString(), "10") }
