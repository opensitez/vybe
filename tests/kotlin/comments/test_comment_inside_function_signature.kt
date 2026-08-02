// vybe-test: kotlin/comments/test_comment_inside_function_signature
// origin: languages/kotlin/tests/kotlin/test_comments.rs

fun total( // comment in signature
            a: Int,
            b: Int
        ): Int = a + b

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((total(2, 3)).toString(), "5")
        }
