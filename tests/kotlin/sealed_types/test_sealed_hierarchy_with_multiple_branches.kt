// vybe-test: kotlin/sealed_types/test_sealed_hierarchy_with_multiple_branches
// origin: languages/kotlin/tests/kotlin/test_sealed_types.rs

sealed class Token {
            class A : Token()
            class B : Token()
            class C : Token()
        }

        fun score(token: Token): Int {
            return when (token) {
                is Token.A -> 1
                is Token.B -> 2
                is Token.C -> 3
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((score(Token.A())).toString(), "1")
            __check((score(Token.B())).toString(), "2")
            __check((score(Token.C())).toString(), "3")
        }
