// vybe-test: kotlin/comments/test_comment_between_when_branches
// origin: languages/kotlin/tests/kotlin/test_comments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = when (1) {
                1 -> 10 // first
                2 -> 20
                else -> 30
            }
            __check((out).toString(), "10")
        }
