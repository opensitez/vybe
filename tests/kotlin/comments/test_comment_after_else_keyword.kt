// vybe-test: kotlin/comments/test_comment_after_else_keyword
// origin: languages/kotlin/tests/kotlin/test_comments.rs

fun main() {
            if (false) {
                println(0)
            } else { // alternate branch
                println(1)
            }
        }

