// vybe-test: kotlin/arrays_ops/test_int_array_fill_with_range
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = intArrayOf(1, 1, 1, 1, 1)
            nums.fill(9, 1, 4)
            __check((nums.joinToString(",")).toString(), "1,9,9,9,1")
        }
