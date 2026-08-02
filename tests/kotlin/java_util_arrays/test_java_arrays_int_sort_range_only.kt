// vybe-test: kotlin/java_util_arrays/test_java_arrays_int_sort_range_only
// origin: languages/kotlin/tests/kotlin/test_java_util_arrays.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val data = intArrayOf(9, 8, 7, 6, 5, 4)
            java.util.Arrays.sort(data, 1, 4)
            __check((java.util.Arrays.toString(data)).toString(), "[9, 6, 7, 8, 5, 4]")
        }
