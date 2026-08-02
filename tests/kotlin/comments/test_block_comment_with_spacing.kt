// vybe-test: kotlin/comments/test_block_comment_with_spacing
// origin: languages/kotlin/tests/kotlin/test_comments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val y = (1 /*a*/) + (2 /*b*/)
            __check((y).toString(), "3")
        }
