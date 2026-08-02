// vybe-test: kotlin/kotlin_bytes_encoding/test_string_from_ascii_bytes_empty
// origin: languages/kotlin/tests/kotlin/test_kotlin_bytes_encoding.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val bytes = byteArrayOf()
            val value = String(bytes)
            __check((bytes.size).toString(), "0")
            __check((value.isEmpty()).toString(), "true")
        }
