// vybe-test: kotlin/java_util_arrays/test_java_arrays_deep_hash_code_is_stable_for_same_structure
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
            __check((java.util.Arrays.deepHashCode(lhs) == java.util.Arrays.deepHashCode(rhs)).toString(), "true")
        }
