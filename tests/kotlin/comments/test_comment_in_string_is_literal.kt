// vybe-test: kotlin/comments/test_comment_in_string_is_literal
// origin: languages/kotlin/tests/kotlin/test_comments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = "// not a comment"
            val other = "a /* block */ b"
            __check((text).toString(), "// not a comment")
            __check((other).toString(), "a /* block */ b")
        }
