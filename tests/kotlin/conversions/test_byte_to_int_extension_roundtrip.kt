// vybe-test: kotlin/conversions/test_byte_to_int_extension_roundtrip
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source: Byte = 13
            val text = source.toString()
            val round = text.toInt()
            __check((source).toString(), "13")
            __check((round).toString(), "13")
            __check((round.toByte()).toString(), "13")
        }
