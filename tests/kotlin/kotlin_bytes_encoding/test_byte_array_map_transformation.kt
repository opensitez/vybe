// vybe-test: kotlin/kotlin_bytes_encoding/test_byte_array_map_transformation
// origin: languages/kotlin/tests/kotlin/test_kotlin_bytes_encoding.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = byteArrayOf(1, 2, 3)
            val doubled = values.map { it * 2 }
            __check((doubled.joinToString(",")).toString(), "2,4,6")
            __check((doubled.sum()).toString(), "12")
        }
