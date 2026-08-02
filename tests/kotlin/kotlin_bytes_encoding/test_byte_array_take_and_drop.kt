// vybe-test: kotlin/kotlin_bytes_encoding/test_byte_array_take_and_drop
// origin: languages/kotlin/tests/kotlin/test_kotlin_bytes_encoding.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = byteArrayOf(1, 2, 3, 4)
            __check((values.take(3).joinToString(",")).toString(), "1,2,3")
            __check((values.takeLast(2).joinToString(",")).toString(), "3,4")
            __check((values.drop(1).joinToString(",")).toString(), "2,3,4")
            __check((values.dropLast(1).joinToString(",")).toString(), "1,2,3")
        }
