// vybe-test: kotlin/destructuring/test_destructuring_in_function_arguments
// origin: languages/kotlin/tests/kotlin/test_destructuring.rs

fun sumPair(pair: Pair<Int, Int>): Int {
            val (left, right) = pair
            return left + right
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((sumPair(Pair(10, 15))).toString(), "25")
        }
