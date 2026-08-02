// vybe-test: kotlin/numeric_types/test_int_boundary_arithmetic_wraps_with_two_complement
// origin: languages/kotlin/tests/kotlin/test_numeric_types.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Int.MAX_VALUE + 1).toString(), "-2147483648")
            __check((Int.MIN_VALUE - 1).toString(), "2147483647")
            __check((Long.MAX_VALUE + 1).toString(), "-9223372036854775808")
            __check((Long.MIN_VALUE - 1).toString(), "9223372036854775807")
        }
