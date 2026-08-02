// vybe-test: kotlin/recursion/test_recursion_depth_guard
// origin: languages/kotlin/tests/kotlin/test_recursion.rs

fun depth(v: Int): Int {
            if (v <= 0) return 0
            return if (v == 1) 1 else 1 + depth(v - 1)
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((depth(1)).toString(), "1")
            __check((depth(4)).toString(), "4")
        }
