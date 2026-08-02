// vybe-test: kotlin/java_util_arrays/test_java_arrays_int_copy_of_shrink
// origin: languages/kotlin/tests/kotlin/test_java_util_arrays.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val data = intArrayOf(1, 2, 3)
            val shrunk = java.util.Arrays.copyOf(data, 2)
            __check((java.util.Arrays.toString(shrunk)).toString(), "[1, 2]")
        }
