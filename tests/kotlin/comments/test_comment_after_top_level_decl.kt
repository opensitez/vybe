// vybe-test: kotlin/comments/test_comment_after_top_level_decl
// origin: languages/kotlin/tests/kotlin/test_comments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val prefix = 2
            val suffix = 3
            __check((prefix * suffix).toString(), "6")
        }
