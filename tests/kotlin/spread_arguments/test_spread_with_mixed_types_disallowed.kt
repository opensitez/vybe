// vybe-test: kotlin/spread_arguments/test_spread_with_mixed_types_disallowed
// origin: languages/kotlin/tests/kotlin/test_spread_arguments.rs

fun collect(vararg values: String): String = values.joinToString(":")
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = arrayOf("x", "y")
            __check((collect(*a)).toString(), "x:y")
        }
