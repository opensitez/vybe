// vybe-test: kotlin/comments/test_comment_in_run_block
// origin: languages/kotlin/tests/kotlin/test_comments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x = run {
                // compute in block
                5
            }
            __check((x).toString(), "5")
        }
