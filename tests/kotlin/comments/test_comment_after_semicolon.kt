// vybe-test: kotlin/comments/test_comment_after_semicolon
// origin: languages/kotlin/tests/kotlin/test_comments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = 10
val b = 5 // trailing comment
            __check((a + b).toString(), "15")
        }
