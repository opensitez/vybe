// vybe-test: kotlin/numeric_types/test_float_to_int_and_to_long_truncate_toward_zero
// origin: languages/kotlin/tests/kotlin/test_numeric_types.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((2.9.toInt()).toString(), "2")
            __check((-2.9.toInt()).toString(), "-2")
            __check((2.9.toLong()).toString(), "2")
            __check((-2.9.toLong()).toString(), "-2")
        }
