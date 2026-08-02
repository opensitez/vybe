// vybe-test: kotlin/arrays_ops/test_int_array_basic_indexing_and_size
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = intArrayOf(1, 2, 3, 4)
            __check((nums.size).toString(), "4")
            __check((nums[0]).toString(), "1")
            __check((nums[3]).toString(), "4")
        }
