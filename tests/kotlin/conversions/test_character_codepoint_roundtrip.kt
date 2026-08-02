// vybe-test: kotlin/conversions/test_character_codepoint_roundtrip
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val ch = 'A'
            val text = ch.toString()
            __check((text).toString(), "A")
            __check((ch).toString(), "A")
        }
