// vybe-test: kotlin/java_util_arrays/test_java_arrays_int_hash_code
// origin: languages/kotlin/tests/kotlin/test_java_util_arrays.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val data = intArrayOf(1, 2, 3)
            __check((java.util.Arrays.hashCode(data)).toString(), "30817")
        }
