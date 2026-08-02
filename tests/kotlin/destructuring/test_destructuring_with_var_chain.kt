// vybe-test: kotlin/destructuring/test_destructuring_with_var_chain
// origin: languages/kotlin/tests/kotlin/test_destructuring.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var (left, right) = Pair(100, 200)
            left += 10
            right -= 20
            val result = left + right
            __check((result).toString(), "290")
        }
