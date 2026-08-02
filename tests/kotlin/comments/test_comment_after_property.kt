// vybe-test: kotlin/comments/test_comment_after_property
// origin: languages/kotlin/tests/kotlin/test_comments.rs

class Box {
            val value = 11 // stored value
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Box().value).toString(), "11")
        }
