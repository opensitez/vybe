// vybe-test: kotlin/comments/test_comment_between_properties
// origin: languages/kotlin/tests/kotlin/test_comments.rs

class Holder {
            val first = 1
            // comment line
            val second = 2
            val third = 3
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val h = Holder()
            __check((h.first + h.second + h.third).toString(), "6")
        }
