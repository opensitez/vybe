// vybe-test: kotlin/characters/test_ascii_codepoint_properties
// origin: languages/kotlin/tests/kotlin/test_characters.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(('A'.code).toString(), "65")
            __check(('z'.code).toString(), "122")
            __check(('0'.code).toString(), "48")
        }
