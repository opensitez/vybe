// vybe-test: kotlin/spread_arguments/test_spread_with_no_varargs
// origin: languages/kotlin/tests/kotlin/test_spread_arguments.rs

fun join(prefix: String, vararg values: Int): String {
            return if (values.isEmpty()) "empty" else prefix + values.joinToString(";")
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((join("x")).toString(), "empty")
        }
