// vybe-test: kotlin/comments/test_comment_next_to_operators
// origin: languages/kotlin/tests/kotlin/test_comments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = 8/*c*/+4
            val b = 2/*c*/+3
            __check((a - b).toString(), "7")
        }
