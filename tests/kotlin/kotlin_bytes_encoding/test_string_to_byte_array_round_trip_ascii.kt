// vybe-test: kotlin/kotlin_bytes_encoding/test_string_to_byte_array_round_trip_ascii
// origin: languages/kotlin/tests/kotlin/test_kotlin_bytes_encoding.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = "Kotlin"
            val bytes = source.toByteArray()
            val value = String(bytes)
            __check((bytes.size).toString(), "6")
            __check((value).toString(), "Kotlin")
        }
