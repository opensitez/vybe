// vybe-test: kotlin/kotlin_numeric_boundary_checks/test_int_maximum_overflow_wraps_like_runtime
// origin: languages/kotlin/tests/kotlin/test_kotlin_numeric_boundary_checks.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Int.MAX_VALUE + 1).toString(), "-2147483648")
        }
