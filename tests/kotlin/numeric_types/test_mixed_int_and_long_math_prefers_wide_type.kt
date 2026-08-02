// vybe-test: kotlin/numeric_types/test_mixed_int_and_long_math_prefers_wide_type
// origin: languages/kotlin/tests/kotlin/test_numeric_types.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base = 3
            val wide = 10L
            __check((base + wide).toString(), "13")
            __check((base * wide).toString(), "30")
            __check((wide - base).toString(), "7")
            __check((wide / base).toString(), "3")
        }
