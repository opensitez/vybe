// vybe-test: kotlin/characters/test_character_combine_case_predicate
// origin: languages/kotlin/tests/kotlin/test_characters.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = "aB3!"
            val lowered = value.map { it.lowercaseChar() }.joinToString("")
            val uppered = value.map { it.uppercaseChar() }.joinToString("")
            __check((lowered).toString(), "ab3!")
            __check((uppered).toString(), "AB3!")
            __check((lowered[1].isUpperCase()).toString(), "false")
            __check((uppered[2].isDigit()).toString(), "true")
        }
