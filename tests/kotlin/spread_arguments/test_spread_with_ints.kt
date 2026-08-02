// vybe-test: kotlin/spread_arguments/test_spread_with_ints
// origin: languages/kotlin/tests/kotlin/test_spread_arguments.rs

fun join(prefix: String, vararg values: Int): String {
            return prefix + values.joinToString(",")
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = intArrayOf(1, 2, 3)
            __check((join("v", *nums)).toString(), "v1,2,3")
        }
