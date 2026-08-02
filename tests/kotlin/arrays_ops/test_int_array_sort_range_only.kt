// vybe-test: kotlin/arrays_ops/test_int_array_sort_range_only
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = intArrayOf(9, 1, 8, 3, 2, 7)
            nums.sort(1, 5)
            __check((nums.joinToString(",")).toString(), "9,1,2,3,8,7")
        }
