// vybe-test: kotlin/conversions/test_string_to_double_or_null_rejects_bad_and_nan_variants
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("3.5".toDoubleOrNull() ?: -1.0).toString(), "3.5")
            __check(("bad".toDoubleOrNull() ?: -1.0).toString(), "-1")
            __check(("NaN".toDoubleOrNull()?.isNaN() ?: false).toString(), "true")
        }
