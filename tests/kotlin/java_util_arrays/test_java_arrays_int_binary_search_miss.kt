// vybe-test: kotlin/java_util_arrays/test_java_arrays_int_binary_search_miss
// origin: languages/kotlin/tests/kotlin/test_java_util_arrays.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val data = intArrayOf(1, 2, 3, 4, 5)
            __check((java.util.Arrays.binarySearch(data, 4)).toString(), "3")
            __check((java.util.Arrays.binarySearch(data, 6)).toString(), "-6")
        }
