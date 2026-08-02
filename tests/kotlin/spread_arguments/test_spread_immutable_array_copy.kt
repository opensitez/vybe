// vybe-test: kotlin/spread_arguments/test_spread_immutable_array_copy
// origin: languages/kotlin/tests/kotlin/test_spread_arguments.rs

fun copy(base: Int, values: IntArray): Int {
            return base + values.sum()
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = intArrayOf(2, 4)
            __check((copy(3, values)).toString(), "9")
        }
