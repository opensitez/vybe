// vybe-test: kotlin/recursion/test_recursion_nested_tail_guard
// origin: languages/kotlin/tests/kotlin/test_recursion.rs

fun nested(n: Int): Int {
            if (n == 0) return 1
            if (n == 1) return 1
            return nested(n - 1) + nested(n - 2)
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((nested(6)).toString(), "8")
        }
