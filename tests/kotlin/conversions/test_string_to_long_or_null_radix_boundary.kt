// vybe-test: kotlin/conversions/test_string_to_long_or_null_radix_boundary
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("7fffffff".toLongOrNull(16) ?: 0).toString(), "2147483647")
            __check(("xyz".toLongOrNull(16) ?: -1).toString(), "-1")
        }
