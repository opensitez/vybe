// vybe-test: kotlin/comments/test_doc_comment_function
// origin: languages/kotlin/tests/kotlin/test_comments.rs

/** doc comment */
        fun add(a: Int, b: Int) = a + b

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((add(1, 2)).toString(), "3")
        }
