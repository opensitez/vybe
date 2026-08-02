// vybe-test: kotlin/destructuring/test_destructuring_with_assignment
// origin: languages/kotlin/tests/kotlin/test_destructuring.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { var (left, right) = Pair(12, 3)
left = left - 2
__check((left * right).toString(), "30") }
