// vybe-test: kotlin/comments/test_comment_in_while_loop
// origin: languages/kotlin/tests/kotlin/test_comments.rs

fun main() {
            var sum = 0
            var i = 0
            while (i < 2) {
                sum += i
                // keep loop
                i += 1
            }
            println(sum)
        }

