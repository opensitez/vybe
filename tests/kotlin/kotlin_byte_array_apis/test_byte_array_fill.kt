// vybe-test: kotlin/kotlin_byte_array_apis/test_byte_array_fill
// origin: languages/kotlin/tests/kotlin/test_kotlin_byte_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = byteArrayOf(1, 1, 1)
            java.util.Arrays.fill(a, 3)
            __check((a[0].toString()).toString(), "3")
            __check((a[2].toString()).toString(), "3")
        }
