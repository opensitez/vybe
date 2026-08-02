// vybe-test: kotlin/kotlin_bytes_encoding/test_byte_array_mutable_set_access
// origin: languages/kotlin/tests/kotlin/test_kotlin_bytes_encoding.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val bytes = byteArrayOf(10, 20, 30)
            bytes[1] = 55
            __check((bytes[0]).toString(), "10")
            __check((bytes[1]).toString(), "55")
            __check((bytes[2]).toString(), "30")
            __check((bytes.joinToString(",")).toString(), "10,55,30")
        }
