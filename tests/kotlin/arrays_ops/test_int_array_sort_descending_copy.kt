// vybe-test: kotlin/arrays_ops/test_int_array_sort_descending_copy
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = intArrayOf(3, 1, 4, 2)
            val desc = nums.sortedArrayDescending()
            __check((desc.joinToString(",")).toString(), "4,3,2,1")
        }
