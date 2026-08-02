// vybe-test: kotlin/comments/test_comment_after_if_condition
// origin: languages/kotlin/tests/kotlin/test_comments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            if (true) // condition comment
            {
                __check((1).toString(), "1")
            }
        }
