// vybe-test: kotlin/comments/test_comment_after_value_expression
// origin: languages/kotlin/tests/kotlin/test_comments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x = 1 + // inline comment
            2
            __check((x).toString(), "3")
        }
