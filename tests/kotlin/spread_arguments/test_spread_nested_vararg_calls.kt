// vybe-test: kotlin/spread_arguments/test_spread_nested_vararg_calls
// origin: languages/kotlin/tests/kotlin/test_spread_arguments.rs

fun outer(prefix: String, vararg values: String): String = prefix + values.joinToString(":")
        fun build(base: Array<String>): String {
            return outer("p", *base)
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((build(arrayOf("x", "y"))).toString(), "px:y")
        }
