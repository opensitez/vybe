// vybe-test: kotlin/scope/test_member_and_receiver_scope_split
// origin: languages/kotlin/tests/kotlin/test_scope.rs

class ScopeProbe {
            val value = 2
            fun combine(): Int {
                val value = 4
                return this.value + value
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((ScopeProbe().combine()).toString(), "6")
        }
