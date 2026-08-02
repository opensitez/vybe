// vybe-test: kotlin/kotlin_string_interpolation_advanced/test_interpolation_of_char_code_and_codepoint
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_interpolation_advanced.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val ch = 'A'
            __check(("${'$'}ch-${'$'}{ch.code}").toString(), "A-65")
        }
