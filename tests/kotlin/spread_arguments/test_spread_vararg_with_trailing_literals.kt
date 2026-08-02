// vybe-test: kotlin/spread_arguments/test_spread_vararg_with_trailing_literals
// origin: languages/kotlin/tests/kotlin/test_spread_arguments.rs

fun join(prefix: String, vararg values: String): String = prefix + values.joinToString(".")
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val mid = arrayOf("b", "c")
            __check((join("a", *mid, "d")).toString(), "a.b.c.d")
        }
