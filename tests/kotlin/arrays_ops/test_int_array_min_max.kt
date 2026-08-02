// vybe-test: kotlin/arrays_ops/test_int_array_min_max
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = intArrayOf(9, 1, 4, -2, 8)
            __check((nums.minOrNull()).toString(), "-2")
            __check((nums.maxOrNull()).toString(), "9")
        }
