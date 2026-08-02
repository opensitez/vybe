// vybe-test: kotlin/comments/test_comment_inside_block_run
// origin: languages/kotlin/tests/kotlin/test_comments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val result = run {
                /*prepare*/
                4
            }
            __check((result).toString(), "4")
        }
