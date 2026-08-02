// vybe-test: kotlin/conversions/test_to_byte_wraps_lower_bit_slice
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((127.toByte().toInt()).toString(), "127")
            __check((128.toByte().toInt()).toString(), "-128")
            __check((255.toByte().toInt()).toString(), "-1")
            __check(((-129).toByte().toInt()).toString(), "127")
        }
