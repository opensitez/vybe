// vybe-test: kotlin/kotlin_numeric_boundary_checks/test_float_to_int_truncation_direction
// origin: languages/kotlin/tests/kotlin/test_kotlin_numeric_boundary_checks.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((3.9.toInt()).toString(), "3")
            __check(((-3.9).toInt()).toString(), "-3")
        }
