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

        var __buf: String = ""

fun __p(s: String) {
    __buf = __buf + s + "\n"
}

fun __pr(s: String) {
    __buf = __buf + s
}

// The final `println` contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted. Written as two equality
// tests rather than trimming: `String.endsWith` is not implemented in Vybe's
// Kotlin (measured — `"ab\n".endsWith("\n")` throws "undefined is not
// callable"), and a harness that cannot run asserts nothing at all. The cargo
// helper split on "\n" and popped trailing empties, so the two forms were
// equivalent there too.
fun __check(want: String) {
    if (__buf != want && __buf != want + "\n") {
        println("FAIL: want [" + want + "] got [" + __buf + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __p((render(MaybeToken.Present("a"))).toString())
            __p((render(MaybeToken.Missing)).toString())
            __p((render(null)).toString())
        
__check("a\nnone\nnull")
}
