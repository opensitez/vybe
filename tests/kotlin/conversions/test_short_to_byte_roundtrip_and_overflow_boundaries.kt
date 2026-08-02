// vybe-test: kotlin/conversions/test_short_to_byte_roundtrip_and_overflow_boundaries
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a: Short = 32000
            val asByte = a.toByte()
            val restored = asByte.toShort()
            __check((a.toByte().toInt()).toString(), "-96")
            __check((restored).toString(), "-96")
            __check(((-129).toByte().toInt()).toString(), "127")
        }
