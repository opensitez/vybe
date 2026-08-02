// vybe-test: kotlin/spread_arguments/test_spread_char_array_to_string_vararg
// origin: languages/kotlin/tests/kotlin/test_spread_arguments.rs

fun pack(prefix: String, vararg values: Char): String = prefix + values.joinToString("")
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val chars = charArrayOf('h', 'i')
            __check((pack("say:", *chars)).toString(), "say:hi")
        }
