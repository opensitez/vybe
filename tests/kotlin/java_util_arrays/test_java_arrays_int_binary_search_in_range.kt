// vybe-test: kotlin/java_util_arrays/test_java_arrays_int_binary_search_in_range
// origin: languages/kotlin/tests/kotlin/test_java_util_arrays.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val data = intArrayOf(10, 20, 30, 40, 50)
            __check((java.util.Arrays.binarySearch(data, 1, 4, 30)).toString(), "2")
            __check((java.util.Arrays.binarySearch(data, 1, 4, 50)).toString(), "-4")
        }
