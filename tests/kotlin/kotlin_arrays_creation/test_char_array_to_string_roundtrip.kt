// vybe-test: kotlin/kotlin_arrays_creation/test_char_array_to_string_roundtrip
// origin: languages/kotlin/tests/kotlin/test_kotlin_arrays_creation.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val chars = charArrayOf('k', 'o', 't', 'l', 'i', 'n')
            val joined = chars.concatToString()
            __check((joined).toString(), "kotlin")
            __check((joined.toCharArray().size).toString(), "6")
        }
