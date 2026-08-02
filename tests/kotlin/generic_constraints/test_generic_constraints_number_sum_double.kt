// vybe-test: kotlin/generic_constraints/test_generic_constraints_number_sum_double
// origin: languages/kotlin/tests/kotlin/test_generic_constraints.rs

fun <T : Number> sum(a: T, b: T): Int = (a.toDouble() + b.toDouble()).toInt()
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((sum(1.2, 3.9)).toString(), "5")
        }
