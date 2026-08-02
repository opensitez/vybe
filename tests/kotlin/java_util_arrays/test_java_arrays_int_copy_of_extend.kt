// vybe-test: kotlin/java_util_arrays/test_java_arrays_int_copy_of_extend
// origin: languages/kotlin/tests/kotlin/test_java_util_arrays.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val data = intArrayOf(1, 2, 3)
            val extended = java.util.Arrays.copyOf(data, 5)
            __check((java.util.Arrays.toString(extended)).toString(), "[1, 2, 3, 0, 0]")
        }
