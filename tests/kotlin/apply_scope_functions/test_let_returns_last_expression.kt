// vybe-test: kotlin/apply_scope_functions/test_let_returns_last_expression
// origin: languages/kotlin/tests/kotlin/test_apply_scope_functions.rs

fun describe(v: Int): String {
            return v.let { it + 1 }
                .let { it * 2 }
                .toString()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((describe(3)).toString(), "8")
        }
