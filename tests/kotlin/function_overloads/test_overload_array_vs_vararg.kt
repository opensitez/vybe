// vybe-test: kotlin/function_overloads/test_overload_array_vs_vararg
// origin: languages/kotlin/tests/kotlin/test_function_overloads.rs

fun total(values: IntArray): Int = values.sum()
        fun total(a: Int, b: Int): Int = a + b
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((total(1, 2)).toString(), "3")
            __check((total(intArrayOf(1, 2, 3))).toString(), "6")
        }
