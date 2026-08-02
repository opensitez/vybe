// vybe-test: kotlin/kotlin_bytes_encoding/test_byte_array_plus_operator
// origin: languages/kotlin/tests/kotlin/test_kotlin_bytes_encoding.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val first = byteArrayOf(1, 2)
            val second = byteArrayOf(3, 4, 5)
            val both = first + second
            __check((both.joinToString(",")).toString(), "1,2,3,4,5")
            __check((first.size).toString(), "2")
            __check((second.size).toString(), "3")
        }
