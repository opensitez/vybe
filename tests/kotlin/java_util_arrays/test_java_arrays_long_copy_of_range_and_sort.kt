// vybe-test: kotlin/java_util_arrays/test_java_arrays_long_copy_of_range_and_sort
// origin: languages/kotlin/tests/kotlin/test_java_util_arrays.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val data = longArrayOf(9L, 4L, 7L, 1L, 8L, 2L)
            val segment = java.util.Arrays.copyOfRange(data, 1, 5)
            java.util.Arrays.sort(segment)
            __check((java.util.Arrays.toString(segment)).toString(), "[1, 4, 7, 8]")
        }
