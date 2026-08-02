// vybe-test: kotlin/varargs/test_vararg_any_type
// origin: languages/kotlin/tests/kotlin/test_varargs.rs

fun describe(vararg values: Any): String = values.joinToString("|") { it.toString() }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((describe("a", 2, true)).toString(), "a|2|true")
        }
