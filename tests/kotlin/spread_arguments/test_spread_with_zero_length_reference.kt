// vybe-test: kotlin/spread_arguments/test_spread_with_zero_length_reference
// origin: languages/kotlin/tests/kotlin/test_spread_arguments.rs

fun join(vararg values: String): String = values.joinToString(",")
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val empty = arrayOf<String>()
            __check((join(*empty)).toString(), "")
        }
