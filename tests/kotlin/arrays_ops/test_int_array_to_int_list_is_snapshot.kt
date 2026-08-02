// vybe-test: kotlin/arrays_ops/test_int_array_to_int_list_is_snapshot
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = intArrayOf(1, 2, 3)
            val list = nums.toTypedArray().toMutableList()
            list[0] = 9
            nums[1] = 5
            __check((nums.joinToString(",")).toString(), "1,5,3")
            __check((list.joinToString(",")).toString(), "9,2,3")
        }
