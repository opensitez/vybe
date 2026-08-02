// vybe-test: kotlin/conversions/test_character_codepoint_and_reverse_conversion
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val ch = 'Ω'
            val code = ch.code
            __check((code).toString(), "937")
            __check((code.toChar()).toString(), "Ω")
            __check((ch.toString()).toString(), "Ω")
        }
