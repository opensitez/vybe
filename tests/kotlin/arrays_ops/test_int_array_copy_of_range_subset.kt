// vybe-test: kotlin/arrays_ops/test_int_array_copy_of_range_subset
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = intArrayOf(10, 20, 30, 40, 50)
            val mid = nums.copyOfRange(1, 4)
            val empty = nums.copyOfRange(4, 4)
            __check((mid.joinToString(",")).toString(), "20,30,40")
            __check((empty.joinToString(",")).toString(), "")
        }
