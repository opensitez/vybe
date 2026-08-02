// vybe-test: kotlin/java_util_arrays/test_java_arrays_boolean_fill_and_string
// origin: languages/kotlin/tests/kotlin/test_java_util_arrays.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val flags = booleanArrayOf(true, false, true, false)
            java.util.Arrays.fill(flags, 1, 4, true)
            __check((java.util.Arrays.toString(flags)).toString(), "[true, true, true, true]")
        }
