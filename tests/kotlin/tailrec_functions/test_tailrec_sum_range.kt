// vybe-test: kotlin/tailrec_functions/test_tailrec_sum_range
// origin: languages/kotlin/tests/kotlin/test_tailrec_functions.rs

tailrec fun sumRange(n: Int, acc: Int = 0): Int {
            return if (n <= 0) acc else sumRange(n - 1, acc + n)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((sumRange(4)).toString(), "10")
        }
