// vybe-test: kotlin/java_util_arrays/test_java_arrays_int_fill_range
// origin: languages/kotlin/tests/kotlin/test_java_util_arrays.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val data = intArrayOf(0, 0, 0, 0, 0)
            java.util.Arrays.fill(data, 1, 4, 7)
            __check((java.util.Arrays.toString(data)).toString(), "[0, 7, 7, 7, 0]")
        }
