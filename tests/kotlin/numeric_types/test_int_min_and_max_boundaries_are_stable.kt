// vybe-test: kotlin/numeric_types/test_int_min_and_max_boundaries_are_stable
// origin: languages/kotlin/tests/kotlin/test_numeric_types.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Int.MAX_VALUE).toString(), "2147483647")
            __check((Int.MIN_VALUE).toString(), "-2147483648")
            __check((Long.MAX_VALUE).toString(), "9223372036854775807")
            __check((Long.MIN_VALUE).toString(), "-9223372036854775808")
        }
