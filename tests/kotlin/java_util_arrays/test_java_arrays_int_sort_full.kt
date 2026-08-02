// vybe-test: kotlin/java_util_arrays/test_java_arrays_int_sort_full
// origin: languages/kotlin/tests/kotlin/test_java_util_arrays.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val data = intArrayOf(5, 1, 4, 3, 2)
            java.util.Arrays.sort(data)
            __check((java.util.Arrays.toString(data)).toString(), "[1, 2, 3, 4, 5]")
        }
