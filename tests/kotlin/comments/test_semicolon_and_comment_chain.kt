// vybe-test: kotlin/comments/test_semicolon_and_comment_chain
// origin: languages/kotlin/tests/kotlin/test_comments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = 1
val b = 2
// pair
            __check((a + b).toString(), "3")
        }
