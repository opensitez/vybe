// vybe-test: kotlin/comments/test_comment_before_else_branch
// origin: languages/kotlin/tests/kotlin/test_comments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = if (false)
                0
            else
                2 // else side
            __check((out).toString(), "2")
        }
