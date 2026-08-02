// vybe-test: kotlin/visibility/test_private_members_are_visible_inside_companion_object_methods
// origin: languages/kotlin/tests/kotlin/test_visibility.rs

class Token {
            private val secret = "ok"

            companion object {
                fun reveal(token: Token): String = token.secret
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Token.reveal(Token())).toString(), "ok")
        }
