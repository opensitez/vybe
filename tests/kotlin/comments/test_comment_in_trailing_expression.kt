// vybe-test: kotlin/comments/test_comment_in_trailing_expression
// origin: languages/kotlin/tests/kotlin/test_comments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = (1 + 2) // addition
            __check((value).toString(), "3")
        }
