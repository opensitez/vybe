// vybe-test: kotlin/kotlin_bytes_encoding/test_byte_array_zip_addition
// origin: languages/kotlin/tests/kotlin/test_kotlin_bytes_encoding.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val left = byteArrayOf(1, 2, 3)
            val right = byteArrayOf(4, 5, 6)
            val zipped = left.zip(right) { a, b -> a + b }
            __check((zipped.joinToString(",")).toString(), "5,7,9")
        }
