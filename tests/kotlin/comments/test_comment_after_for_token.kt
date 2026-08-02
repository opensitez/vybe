// vybe-test: kotlin/comments/test_comment_after_for_token
// origin: languages/kotlin/tests/kotlin/test_comments.rs

fun main() {
            val nums = listOf(1, 2, 3)
            var sum = 0
            for (n in nums) { // iterate
                sum += n
            }
            println(sum)
        }

