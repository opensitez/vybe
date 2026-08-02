// vybe-test: kotlin/sealed_types/test_sealed_when_expression_with_else_keeps_runtime_branch
// origin: languages/kotlin/tests/kotlin/test_sealed_types.rs

sealed class Token {
            class Left : Token()
            class Right : Token()
        }

        fun describe(token: Token): String {
            return when (token) {
                is Token.Left -> "left"
                is Token.Right -> "right"
                else -> "other"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((describe(Token.Left())).toString(), "left")
            __check((describe(Token.Right())).toString(), "right")
        }
