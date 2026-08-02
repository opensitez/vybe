// vybe-test: kotlin/recursion/test_recursion_while_to_tail_style
// origin: languages/kotlin/tests/kotlin/test_recursion.rs

fun countdown(n: Int): Int {
            return if (n == 0) 0 else 1 + countdown(n - 1)
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((countdown(0)).toString(), "0")
            __check((countdown(2)).toString(), "2")
        }
