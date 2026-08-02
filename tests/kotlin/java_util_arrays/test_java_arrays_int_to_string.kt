// vybe-test: kotlin/java_util_arrays/test_java_arrays_int_to_string
// origin: languages/kotlin/tests/kotlin/test_java_util_arrays.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val data = intArrayOf(5, 1, 4, 3, 2)
            __check((java.util.Arrays.toString(data)).toString(), "[5, 1, 4, 3, 2]")
        }
