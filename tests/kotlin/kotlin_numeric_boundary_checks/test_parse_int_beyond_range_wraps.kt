// vybe-test: kotlin/kotlin_numeric_boundary_checks/test_parse_int_beyond_range_wraps
// origin: languages/kotlin/tests/kotlin/test_kotlin_numeric_boundary_checks.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("2147483647".toInt()).toString(), "2147483647")
            __check(("-2147483648".toLong()).toString(), "-2147483648")
        }
