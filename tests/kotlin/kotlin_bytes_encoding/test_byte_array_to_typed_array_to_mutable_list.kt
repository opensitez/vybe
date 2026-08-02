// vybe-test: kotlin/kotlin_bytes_encoding/test_byte_array_to_typed_array_to_mutable_list
// origin: languages/kotlin/tests/kotlin/test_kotlin_bytes_encoding.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = byteArrayOf(1, 2, 3).toTypedArray().toMutableList()
            values.add(4)
            __check((values.joinToString(",")).toString(), "1,2,3,4")
            __check((values.size).toString(), "4")
        }
