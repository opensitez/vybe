// vybe-test: kotlin/conversions/test_to_short_wraps_lower_bits
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((32767.toShort().toInt()).toString(), "32767")
            __check((32768.toShort().toInt()).toString(), "-32768")
            __check(((-32768).toShort().toInt()).toString(), "-32768")
            __check(((-32769).toShort().toInt()).toString(), "32767")
        }
