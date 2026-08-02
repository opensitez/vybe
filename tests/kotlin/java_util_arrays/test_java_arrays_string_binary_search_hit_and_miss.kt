// vybe-test: kotlin/java_util_arrays/test_java_arrays_string_binary_search_hit_and_miss
// origin: languages/kotlin/tests/kotlin/test_java_util_arrays.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val data = arrayOf("a", "b", "m", "z")
            java.util.Arrays.sort(data)
            __check((java.util.Arrays.binarySearch(data, "m")).toString(), "2")
            __check((java.util.Arrays.binarySearch(data, "k")).toString(), "-3")
        }
