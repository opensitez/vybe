// vybe-test: kotlin/characters/test_character_conditional_branching_on_case
// origin: languages/kotlin/tests/kotlin/test_characters.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = 'G'
            val bucket = when {
                c.isUpperCase() -> "upper"
                c.isDigit() -> "digit"
                else -> "other"
            }
            val c2 = '4'
            val bucket2 = when {
                c2.isUpperCase() -> "upper"
                c2.isLowerCase() -> "lower"
                c2.isDigit() -> "digit"
                else -> "other"
            }
            __check((bucket).toString(), "upper")
            __check((bucket2).toString(), "digit")
        }
