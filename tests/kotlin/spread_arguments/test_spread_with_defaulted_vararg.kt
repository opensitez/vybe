// vybe-test: kotlin/spread_arguments/test_spread_with_defaulted_vararg
// origin: languages/kotlin/tests/kotlin/test_spread_arguments.rs

fun prefix(base: String, vararg values: Int = intArrayOf(9)): String {
            return base + values.joinToString(".")
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((prefix("a")).toString(), "a9")
            __check((prefix("a", 1, 2)).toString(), "a1.2")
        }
