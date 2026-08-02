// vybe-test: kotlin/spread_arguments/test_spread_through_vararg_array_param
// origin: languages/kotlin/tests/kotlin/test_spread_arguments.rs

fun flatten(prefix: String, values: IntArray): String = prefix + values.joinToString("|")
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = intArrayOf(1)
            val b = intArrayOf(2, 3)
            __check((flatten("x", intArrayOf(*a, *b))).toString(), "x1|2|3")
        }
