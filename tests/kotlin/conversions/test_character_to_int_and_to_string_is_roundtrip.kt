// vybe-test: kotlin/conversions/test_character_to_int_and_to_string_is_roundtrip
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val ch = 'z'
            val code = ch.code
            val decoded = code.toChar()
            __check((code).toString(), "122")
            __check((decoded).toString(), "z")
        }
