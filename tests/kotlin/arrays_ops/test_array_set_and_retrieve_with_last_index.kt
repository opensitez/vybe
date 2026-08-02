// vybe-test: kotlin/arrays_ops/test_array_set_and_retrieve_with_last_index
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = IntArray(3)
            nums[nums.lastIndex] = 99
            __check((nums.last()).toString(), "99")
            __check((nums.lastIndex).toString(), "2")
        }
