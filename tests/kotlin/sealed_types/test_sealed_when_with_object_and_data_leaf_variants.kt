// vybe-test: kotlin/sealed_types/test_sealed_when_with_object_and_data_leaf_variants
// origin: languages/kotlin/tests/kotlin/test_sealed_types.rs

sealed class Token {
            data class Named(val label: String) : Token()
            class Number(val value: Int) : Token()
            object Idle : Token()
        }

        fun classify(token: Token): String {
            return when (token) {
                is Token.Named -> token.label
                is Token.Number -> token.value.toString()
                is Token.Idle -> "idle"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((classify(Token.Named("ok"))).toString(), "ok")
            __check((classify(Token.Number(7))).toString(), "7")
            __check((classify(Token.Idle)).toString(), "idle")
        }
