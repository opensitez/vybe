// vybe-test: kotlin/comments/test_comment_in_lambda_body
// origin: languages/kotlin/tests/kotlin/test_comments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val f = { x: Int ->
                // lambda body
                x + 1
            }
            __check((f(2)).toString(), "3")
        }
