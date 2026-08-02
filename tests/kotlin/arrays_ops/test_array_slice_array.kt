// vybe-test: kotlin/arrays_ops/test_array_slice_array
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = intArrayOf(10, 20, 30, 40, 50)
            val slice = nums.sliceArray(1..3)
            __check((slice.joinToString(",")).toString(), "20,30,40")
        }
