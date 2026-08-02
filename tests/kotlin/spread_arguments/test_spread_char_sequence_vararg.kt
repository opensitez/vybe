// vybe-test: kotlin/spread_arguments/test_spread_char_sequence_vararg
// origin: languages/kotlin/tests/kotlin/test_spread_arguments.rs

fun concat(vararg items: Char): String {
            return items.joinToString("")
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val head = charArrayOf('x', 'y')
            __check((concat(*head, 'z')).toString(), "xyz")
        }
