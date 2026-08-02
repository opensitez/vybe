// vybe-test: kotlin/comments/test_comment_before_value
// origin: languages/kotlin/tests/kotlin/test_comments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 12
            // compute result
            __check((value).toString(), "12")
        }
