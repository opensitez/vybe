// vybe-test: kotlin/literals/test_character_unicode_literal
// origin: languages/kotlin/tests/kotlin/test_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val omega = '\u03A9'
            val heart = '\u2665'
            __check((omega).toString(), "Ω")
            __check((heart).toString(), "♥")
        }
