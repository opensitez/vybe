// vybe-test: kotlin/comments/test_comment_on_object_member
// origin: languages/kotlin/tests/kotlin/test_comments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val obj = object {
                val one = 1 // member comment
            }
            __check((obj.one).toString(), "1")
        }
