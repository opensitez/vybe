// vybe-test: kotlin/java_util_arrays/test_java_arrays_int_copy_of_range_middle
// origin: languages/kotlin/tests/kotlin/test_java_util_arrays.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val data = intArrayOf(1, 2, 3, 4, 5)
            val segment = java.util.Arrays.copyOfRange(data, 1, 4)
            __check((java.util.Arrays.toString(segment)).toString(), "[2, 3, 4]")
        }
