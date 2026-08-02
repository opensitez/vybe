// vybe-test: kotlin/comments/test_comment_line_separates_statements
// origin: languages/kotlin/tests/kotlin/test_comments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base = 4
            // bump value
            val value = base + 1
            __check((value).toString(), "5")
        }
