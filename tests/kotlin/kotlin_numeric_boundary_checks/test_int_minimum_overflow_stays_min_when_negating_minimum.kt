// vybe-test: kotlin/kotlin_numeric_boundary_checks/test_int_minimum_overflow_stays_min_when_negating_minimum
// origin: languages/kotlin/tests/kotlin/test_kotlin_numeric_boundary_checks.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Int.MIN_VALUE * -1).toString(), "-2147483648")
        }
