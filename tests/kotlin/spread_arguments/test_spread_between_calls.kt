// vybe-test: kotlin/spread_arguments/test_spread_between_calls
// origin: languages/kotlin/tests/kotlin/test_spread_arguments.rs

fun wrap(base: String, values: IntArray): String = base + values.joinToString(",")
        fun sink(prefix: String, vararg values: Int): Int {
            return values.sum() + prefix.length
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = intArrayOf(2, 3, 4)
            __check((wrap("p", values)).toString(), "p2,3,4")
            __check((sink("x", *values)).toString(), "8")
        }
