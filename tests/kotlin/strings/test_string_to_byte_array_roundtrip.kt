// vybe-test: kotlin/strings/test_string_to_byte_array_roundtrip
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val bytes = "ok".toByteArray()
            val back = String(bytes)
            __check((bytes.size).toString(), "2")
            __check((back).toString(), "ok")
        }
