// vybe-test: kotlin/arrays_ops/test_int_array_sum_and_average
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = intArrayOf(1, 2, 3, 4)
            __check((nums.sum()).toString(), "10")
            __check((nums.average()).toString(), "2.5")
        }
