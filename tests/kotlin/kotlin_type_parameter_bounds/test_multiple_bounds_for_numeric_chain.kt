// vybe-test: kotlin/kotlin_type_parameter_bounds/test_multiple_bounds_for_numeric_chain
// origin: languages/kotlin/tests/kotlin/test_kotlin_type_parameter_bounds.rs

fun <T> sumPositive(a: T, b: T): Double where T : Number, T : Comparable<T> {
            return a.toDouble() + b.toDouble()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((sumPositive(2, 3)).toString(), "5")
            __check((sumPositive(1.5, 2.5)).toString(), "4")
        }
