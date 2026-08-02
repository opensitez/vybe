// vybe-test: kotlin/conversions/test_string_to_int_or_null_and_numeric_nullability
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("42".toIntOrNull() ?: -1).toString(), "42")
            __check(("nope".toIntOrNull() ?: -1).toString(), "-1")
            __check(("-9".toIntOrNull() ?: 0).toString(), "-9")
        }
