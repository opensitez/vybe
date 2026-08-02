// vybe-test: kotlin/comments/test_comment_between_type_and_initializer
// origin: languages/kotlin/tests/kotlin/test_comments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base: Int // type comment
            = 6
            __check((base).toString(), "6")
        }
