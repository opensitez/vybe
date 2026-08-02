// vybe-test: kotlin/literals/test_char_and_string_literal_mixing
// origin: languages/kotlin/tests/kotlin/test_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c: Char = 'x'
            val text = "c=$c"
            val pair = listOf(c, 'y')
            __check((text).toString(), "c=x")
            __check((pair.joinToString("-")).toString(), "x-y")
        }
