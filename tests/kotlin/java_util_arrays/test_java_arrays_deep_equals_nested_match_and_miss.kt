// vybe-test: kotlin/java_util_arrays/test_java_arrays_deep_equals_nested_match_and_miss
// origin: languages/kotlin/tests/kotlin/test_java_util_arrays.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val lhs = arrayOf(arrayOf(1, 2), arrayOf(3, 4))
            val rhs = arrayOf(arrayOf(1, 2), arrayOf(3, 4))
            val other = arrayOf(arrayOf(1, 2), arrayOf(3, 5))
            __check((java.util.Arrays.deepEquals(lhs, rhs)).toString(), "true")
            __check((java.util.Arrays.deepEquals(lhs, other)).toString(), "false")
        }
