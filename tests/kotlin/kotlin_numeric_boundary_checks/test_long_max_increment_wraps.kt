// vybe-test: kotlin/kotlin_numeric_boundary_checks/test_long_max_increment_wraps
// origin: languages/kotlin/tests/kotlin/test_kotlin_numeric_boundary_checks.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Long.MAX_VALUE + 1).toString(), "-9223372036854775808")
        }
