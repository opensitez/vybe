// vybe-test: kotlin/recursion/test_recursion_tail_ish_sum
// origin: languages/kotlin/tests/kotlin/test_recursion.rs

fun sumTail(n: Int, acc: Int = 0): Int {
            return if (n <= 0) acc else sumTail(n - 1, acc + n)
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((sumTail(4)).toString(), "10")
        }
