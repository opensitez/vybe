// vybe-test: kotlin/infix/test_infix_character_range_membership
// origin: languages/kotlin/tests/kotlin/test_infix.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(('b' in 'a'..'c').toString(), "true")
            __check(('z' in 'a'..'c').toString(), "false")
        }
