// vybe-test: kotlin/java_util_arrays/test_java_arrays_string_copy_of_range
// origin: languages/kotlin/tests/kotlin/test_java_util_arrays.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = arrayOf("a", "b", "c", "d")
            val segment = java.util.Arrays.copyOfRange(source, 2, 4)
            __check((java.util.Arrays.toString(segment)).toString(), "[c, d]")
        }
