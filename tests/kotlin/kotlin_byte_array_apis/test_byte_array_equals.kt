// vybe-test: kotlin/kotlin_byte_array_apis/test_byte_array_equals
// origin: languages/kotlin/tests/kotlin/test_kotlin_byte_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = byteArrayOf(1, 2)
            val b = byteArrayOf(1, 2)
            __check(((a == b).toString()).toString(), "false")
        }
