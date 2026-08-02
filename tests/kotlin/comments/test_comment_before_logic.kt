// vybe-test: kotlin/comments/test_comment_before_logic
// origin: languages/kotlin/tests/kotlin/test_comments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val first = 3
            val second = 4
            // branch comment
            __check((first + second).toString(), "7")
        }
