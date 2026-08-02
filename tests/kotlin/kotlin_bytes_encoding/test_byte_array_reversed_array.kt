// vybe-test: kotlin/kotlin_bytes_encoding/test_byte_array_reversed_array
// origin: languages/kotlin/tests/kotlin/test_kotlin_bytes_encoding.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = byteArrayOf(1, 3, 5)
            val reversed = values.reversedArray()
            __check((reversed.joinToString(",")).toString(), "5,3,1")
            __check((values.joinToString(",")).toString(), "1,3,5")
        }
