// vybe-test: kotlin/comments/test_block_comment_in_arithmetic
// origin: languages/kotlin/tests/kotlin/test_comments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x = 1 /* comment */ + 2
            __check((x).toString(), "3")
        }
