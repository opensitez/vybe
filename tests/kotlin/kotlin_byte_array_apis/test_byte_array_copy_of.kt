// vybe-test: kotlin/kotlin_byte_array_apis/test_byte_array_copy_of
// origin: languages/kotlin/tests/kotlin/test_kotlin_byte_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = byteArrayOf(1, 2)
            val b = a.copyOf()
            b[0] = 8
            __check((a[0].toString()).toString(), "1")
            __check((b[0].toString()).toString(), "8")
        }
