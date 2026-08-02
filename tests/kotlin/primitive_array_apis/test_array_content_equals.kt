// vybe-test: kotlin/primitive_array_apis/test_array_content_equals
// origin: languages/kotlin/tests/kotlin/test_primitive_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = intArrayOf(1, 2)
            val b = intArrayOf(1, 2)
            __check((a.contentEquals(b)).toString(), "true")
            __check((a.contentHashCode() == b.contentHashCode()).toString(), "true")
        }
