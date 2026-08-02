// vybe-test: kotlin/kotlin_number_conversion_apis/test_number_to_byte_and_short
// origin: languages/kotlin/tests/kotlin/test_kotlin_number_conversion_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((1234.7.toInt()).toString(), "1234")
            __check((129.toByte()).toString(), "-126")
            __check((128.toByte()).toString(), "-128")
            __check((32000.toShort()).toString(), "32000")
            __check((40000.toShort()).toString(), "-25536")
        }
