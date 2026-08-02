// vybe-test: kotlin/spread_arguments/test_spread_in_lambda_capture
// origin: languages/kotlin/tests/kotlin/test_spread_arguments.rs

fun total(vararg values: Int): Int = values.sum()
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = intArrayOf(5, 6)
            val runTotal = { arr: IntArray -> total(*arr) }
            __check((runTotal(source)).toString(), "11")
        }
