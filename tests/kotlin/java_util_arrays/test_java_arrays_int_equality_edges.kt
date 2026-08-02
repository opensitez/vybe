// vybe-test: kotlin/java_util_arrays/test_java_arrays_int_equality_edges
// origin: languages/kotlin/tests/kotlin/test_java_util_arrays.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = intArrayOf(1, 2, 3)
            val b = intArrayOf(1, 2, 3)
            val c = intArrayOf(1, 2, 4)
            __check((java.util.Arrays.equals(a, b)).toString(), "true")
            __check((java.util.Arrays.equals(a, c)).toString(), "false")
        }
