// vybe-test: kotlin/step_ranges/test_char_range
// origin: languages/kotlin/tests/kotlin/test_step_ranges.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = ('a'..'e').toList()
            __check((values.joinToString("")).toString(), "abcde")
            __check((('e' downTo 'c').toList().joinToString("")).toString(), "edc")
        }
