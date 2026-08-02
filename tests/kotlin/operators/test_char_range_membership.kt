// vybe-test: kotlin/operators/test_char_range_membership
// origin: languages/kotlin/tests/kotlin/test_operators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val vowels = 'a'..'f'
            __check(('c' in vowels).toString(), "true")
            __check(('z' in vowels).toString(), "false")
            __check(('a' !in vowels).toString(), "false")
        }
