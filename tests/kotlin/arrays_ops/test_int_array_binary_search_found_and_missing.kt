// vybe-test: kotlin/arrays_ops/test_int_array_binary_search_found_and_missing
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = intArrayOf(1, 3, 5, 7)
            __check((nums.binarySearch(5)).toString(), "2")
            __check((nums.binarySearch(4)).toString(), "-3")
        }
