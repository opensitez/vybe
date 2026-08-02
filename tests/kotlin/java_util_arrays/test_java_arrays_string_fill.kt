// vybe-test: kotlin/java_util_arrays/test_java_arrays_string_fill
// origin: languages/kotlin/tests/kotlin/test_java_util_arrays.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val data = arrayOf("left", "right", "up")
            java.util.Arrays.fill(data, 1, 2, "mid")
            __check((java.util.Arrays.toString(data)).toString(), "[left, mid, up]")
        }
