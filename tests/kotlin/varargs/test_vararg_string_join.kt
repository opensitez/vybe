// vybe-test: kotlin/varargs/test_vararg_string_join
// origin: languages/kotlin/tests/kotlin/test_varargs.rs

fun join(base: String, vararg values: String): String =
            base + values.joinToString(":")

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((join("x", "a", "b", "c")).toString(), "xa:b:c")
        }
