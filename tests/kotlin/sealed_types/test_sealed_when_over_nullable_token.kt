// vybe-test: kotlin/sealed_types/test_sealed_when_over_nullable_token
// origin: languages/kotlin/tests/kotlin/test_sealed_types.rs

sealed class MaybeToken {
            class Present(val value: String) : MaybeToken()
            object Missing : MaybeToken()
        }

        fun render(token: MaybeToken?): String {
            return when (token) {
                is MaybeToken.Present -> token.value
                is MaybeToken.Missing -> "none"
                null -> "null"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((render(MaybeToken.Present("a"))).toString(), "a")
            __check((render(MaybeToken.Missing)).toString(), "none")
            __check((render(null)).toString(), "null")
        }
