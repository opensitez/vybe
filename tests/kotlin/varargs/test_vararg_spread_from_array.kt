// vybe-test: kotlin/varargs/test_vararg_spread_from_array
// origin: languages/kotlin/tests/kotlin/test_varargs.rs

fun maxOfAll(vararg values: Int): Int = values.maxOrNull() ?: 0

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base = intArrayOf(4, 1, 8)
            __check((maxOfAll(*base)).toString(), "8")
        }
