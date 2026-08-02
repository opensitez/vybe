// vybe-test: kotlin/recursion/test_recursion_range_step
// origin: languages/kotlin/tests/kotlin/test_recursion.rs

fun countRange(start: Int, end: Int): Int = if (start > end) 0 else 1 + countRange(start + 1, end)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((countRange(1, 3)).toString(), "3")
        }
