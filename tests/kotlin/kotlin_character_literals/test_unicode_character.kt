// vybe-test: kotlin/kotlin_character_literals/test_unicode_character
// origin: languages/kotlin/tests/kotlin/test_kotlin_character_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val g = '\u0047'
            val omega = '\u03A9'
            __check((g).toString(), "G")
            __check((omega).toString(), "Ω")
        }
