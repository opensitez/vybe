// vybe-test: kotlin/conversions/test_long_to_int_overflow_wraps
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val large = 3_000_000_000L
            __check((large.toInt()).toString(), "-1294967296")
        }
