// vybe-test: kotlin/function_overloads/test_overload_with_defaulted_tail_param
// origin: languages/kotlin/tests/kotlin/test_function_overloads.rs

fun pair(a: Int): String = "single" + a
        fun pair(a: Int, b: Int = 1): String = "pair" + (a + b)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((pair(3)).toString(), "single3")
            __check((pair(3, 2)).toString(), "pair5")
        }
