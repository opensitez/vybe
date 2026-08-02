// vybe-test: kotlin/arrays_ops/test_array_binary_search_with_comparator_like_transform
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = intArrayOf(1, 3, 5, 7, 9)
            __check((nums.binarySearch(5)).toString(), "2")
            __check((nums.binarySearch(6)).toString(), "-4")
            __check((nums.binarySearch(6, 0, 4)).toString(), "-4")
        }
