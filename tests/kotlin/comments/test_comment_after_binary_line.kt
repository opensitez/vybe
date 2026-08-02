// vybe-test: kotlin/comments/test_comment_after_binary_line
// origin: languages/kotlin/tests/kotlin/test_comments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val total = 10 + // plus
                20 + // plus
                30
            __check((total).toString(), "60")
        }
