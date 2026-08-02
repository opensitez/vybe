// vybe-test: kotlin/java_util_arrays/test_java_arrays_char_fill_and_sort
// origin: languages/kotlin/tests/kotlin/test_java_util_arrays.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val data = charArrayOf('z', 'a', 'c', 'b')
            java.util.Arrays.sort(data)
            java.util.Arrays.fill(data, 1, 3, 'x')
            __check((java.util.Arrays.toString(data)).toString(), "[a, x, x, z]")
        }
