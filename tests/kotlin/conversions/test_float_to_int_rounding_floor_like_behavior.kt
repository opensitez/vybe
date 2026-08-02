// vybe-test: kotlin/conversions/test_float_to_int_rounding_floor_like_behavior
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((9.999999.toInt()).toString(), "9")
            __check(((-9.999999).toInt()).toString(), "-9")
        }
