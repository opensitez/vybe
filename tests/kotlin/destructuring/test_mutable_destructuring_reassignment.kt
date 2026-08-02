// vybe-test: kotlin/destructuring/test_mutable_destructuring_reassignment
// origin: languages/kotlin/tests/kotlin/test_destructuring.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var (left, right) = Pair(10, 20)
            left += 1
            val sum = left + right
            __check((sum).toString(), "31")
        }
