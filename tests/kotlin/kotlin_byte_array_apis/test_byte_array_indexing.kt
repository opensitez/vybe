// vybe-test: kotlin/kotlin_byte_array_apis/test_byte_array_indexing
// origin: languages/kotlin/tests/kotlin/test_kotlin_byte_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val data = byteArrayOf(5, 7, 9)
            __check((data[1].toString()).toString(), "7")
        }
