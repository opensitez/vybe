// vybe-test: kotlin/comments/test_comment_between_class_members
// origin: languages/kotlin/tests/kotlin/test_comments.rs

class Pairish {
            val a = 1
            // separator
            val b = 2
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val p = Pairish()
            __check((p.a + p.b).toString(), "3")
        }
