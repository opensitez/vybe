// vybe-test: kotlin/bitwise_operations/test_short_and_byte_are_extended_before_bitwise
// origin: languages/kotlin/tests/kotlin/test_bitwise_operations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val signedByte: Byte = -1
            val signedShort: Short = -2
            val byteUnsigned = signedByte.toInt() and 0xFF
            val shortUnsigned = signedShort.toInt() and 0xFFFF
            val combined = (byteUnsigned and shortUnsigned)
            __check((byteUnsigned).toString(), "255")
            __check((shortUnsigned).toString(), "65534")
            __check((combined).toString(), "254")
        }
