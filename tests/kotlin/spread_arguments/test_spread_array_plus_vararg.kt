// vybe-test: kotlin/spread_arguments/test_spread_array_plus_vararg
// origin: languages/kotlin/tests/kotlin/test_spread_arguments.rs

fun combine(base: String, vararg tags: String): String = base + tags.joinToString("|")
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val tags = arrayOf("a", "b")
            __check((combine("x", "c", *tags, "d")).toString(), "xa|b|d|c")
        }
