// vybe-test: kotlin/kotlin_numeric_boundary_checks/test_double_positive_to_short_is_truncated
// origin: languages/kotlin/tests/kotlin/test_kotlin_numeric_boundary_checks.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((128.9.toShort()).toString(), "-128")
            __check((-129.1.toShort()).toString(), "127")
        }
