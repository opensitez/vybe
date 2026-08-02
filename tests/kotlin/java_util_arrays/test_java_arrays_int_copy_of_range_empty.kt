// vybe-test: kotlin/java_util_arrays/test_java_arrays_int_copy_of_range_empty
// origin: languages/kotlin/tests/kotlin/test_java_util_arrays.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val data = intArrayOf(1, 2, 3)
            val empty = java.util.Arrays.copyOfRange(data, 2, 2)
            __check((java.util.Arrays.toString(empty)).toString(), "[]")
        }
