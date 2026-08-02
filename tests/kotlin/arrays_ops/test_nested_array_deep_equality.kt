// vybe-test: kotlin/arrays_ops/test_nested_array_deep_equality
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = arrayOf(arrayOf(1, 2), arrayOf(3))
            val b = arrayOf(arrayOf(1, 2), arrayOf(3))
            val c = arrayOf(arrayOf(1, 2), arrayOf(4))
            __check((a.contentDeepEquals(b)).toString(), "true")
            __check((a.contentDeepEquals(c)).toString(), "false")
            __check((a.contentDeepToString()).toString(), "[[1, 2], [3]]")
        }
