// vybe-test: kotlin/kotlin_type_parameter_bounds/test_upper_bound_comparable_generic_function
// origin: languages/kotlin/tests/kotlin/test_kotlin_type_parameter_bounds.rs

fun <T : Comparable<T>> maxValue(a: T, b: T): T = if (a >= b) a else b

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((maxValue(3, 7)).toString(), "7")
            __check((maxValue("a", "b")).toString(), "b")
        }
