// vybe-test: kotlin/kotlin_byte_array_apis/test_byte_array_empty
// origin: languages/kotlin/tests/kotlin/test_kotlin_byte_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val empty = byteArrayOf()
            __check((empty.size).toString(), "0")
            __check((empty.isEmpty().toString()).toString(), "true")
        }
