// vybe-test: kotlin/java_util_arrays/test_java_arrays_deep_to_string_nested
// origin: languages/kotlin/tests/kotlin/test_java_util_arrays.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nested = arrayOf(arrayOf(1, 2), arrayOf(3, 4))
            __check((java.util.Arrays.deepToString(nested)).toString(), "[[1, 2], [3, 4]]")
        }
