// vybe-test: kotlin/arrays_ops/test_array_reverse_in_place_and_copy
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = intArrayOf(1, 2, 3, 4)
            nums.reverse()
            __check((nums.joinToString(",")).toString(), "4,3,2,1")
            val back = nums.reversedArray()
            __check((back.joinToString(",")).toString(), "1,2,3,4")
            __check((nums[0] + back[0]).toString(), "5")
        }
