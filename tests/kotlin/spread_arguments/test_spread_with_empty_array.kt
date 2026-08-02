// vybe-test: kotlin/spread_arguments/test_spread_with_empty_array
// origin: languages/kotlin/tests/kotlin/test_spread_arguments.rs

fun join(vararg values: Int): Int = values.size
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val empty = intArrayOf()
            __check((join(*empty)).toString(), "0")
        }
