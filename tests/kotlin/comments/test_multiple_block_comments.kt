// vybe-test: kotlin/comments/test_multiple_block_comments
// origin: languages/kotlin/tests/kotlin/test_comments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 2 + /*a*/ 3 + /*b*/ 4
            __check((value).toString(), "9")
        }
