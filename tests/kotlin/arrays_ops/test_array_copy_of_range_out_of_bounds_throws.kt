// vybe-test: kotlin/arrays_ops/test_array_copy_of_range_out_of_bounds_throws
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = intArrayOf(1, 2, 3)
            try {
                nums.copyOfRange(-1, 2)
            } catch (e: IllegalArgumentException) {
                __check(("bad").toString(), "bad")
            }
        }
