// vybe-test: kotlin/spread_arguments/test_spread_long_array
// origin: languages/kotlin/tests/kotlin/test_spread_arguments.rs

fun joinAll(prefix: String, vararg values: Long): String {
            return prefix + values.joinToString(",")
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = longArrayOf(5L, 6L)
            __check((joinAll("L", *nums)).toString(), "L5,6")
        }
