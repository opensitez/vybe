// vybe-test: kotlin/equality_hashcode/test_array_content_equality_uses_reference_only
// origin: languages/kotlin/tests/kotlin/test_equality_hashcode.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val left = arrayOf(1, 2, 3)
            val right = arrayOf(1, 2, 3)
            __check((left == right).toString(), "false")
            __check((left.contentToString()).toString(), "[1, 2, 3]")
            __check((contentDeepToString(arrayOf(left, right))).toString(), "[[1, 2, 3], [1, 2, 3]]")
        }
