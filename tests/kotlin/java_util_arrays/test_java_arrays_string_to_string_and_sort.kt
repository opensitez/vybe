// vybe-test: kotlin/java_util_arrays/test_java_arrays_string_to_string_and_sort
// origin: languages/kotlin/tests/kotlin/test_java_util_arrays.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val data = arrayOf("k", "a", "m", "b")
            java.util.Arrays.sort(data)
            __check((java.util.Arrays.toString(data)).toString(), "[a, b, k, m]")
        }
