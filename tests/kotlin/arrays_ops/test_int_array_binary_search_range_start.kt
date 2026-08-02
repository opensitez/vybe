// vybe-test: kotlin/arrays_ops/test_int_array_binary_search_range_start
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = intArrayOf(1, 2, 3, 4, 5, 6)
            __check((nums.binarySearch(4, 1, 4)).toString(), "3")
            __check((nums.binarySearch(4, 0, 3)).toString(), "-5")
        }
