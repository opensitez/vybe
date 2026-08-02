// vybe-test: kotlin/comments/test_comment_between_annotations
// origin: languages/kotlin/tests/kotlin/test_comments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val count: Int // typed
            = 9
            __check((count).toString(), "9")
        }
