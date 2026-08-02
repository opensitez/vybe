// vybe-test: kotlin/kotlin_byte_array_apis/test_byte_array_set
// origin: languages/kotlin/tests/kotlin/test_kotlin_byte_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val data = byteArrayOf(1, 2)
            data[0] = 9
            __check((data[0].toString()).toString(), "9")
        }
