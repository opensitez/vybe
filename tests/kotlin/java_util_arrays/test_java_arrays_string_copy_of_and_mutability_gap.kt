// vybe-test: kotlin/java_util_arrays/test_java_arrays_string_copy_of_and_mutability_gap
// origin: languages/kotlin/tests/kotlin/test_java_util_arrays.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = arrayOf("x", "y", "z")
            val copy = java.util.Arrays.copyOf(source, 4)
            copy[1] = "changed"
            __check((source[1]).toString(), "y")
            __check((java.util.Arrays.toString(copy)).toString(), "[x, changed, z, null]")
        }
