// vybe-test: kotlin/arrays_ops/test_array_content_equality_for_object_arrays
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = arrayOf("x", "y")
            val b = arrayOf("x", "y")
            val c = arrayOf("x", "z")
            __check((a.contentEquals(b)).toString(), "true")
            __check((a.contentEquals(c)).toString(), "false")
            __check((a.contentToString()).toString(), "[x, y]")
        }
