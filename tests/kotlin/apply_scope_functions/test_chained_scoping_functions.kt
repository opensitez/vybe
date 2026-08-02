// vybe-test: kotlin/apply_scope_functions/test_chained_scoping_functions
// origin: languages/kotlin/tests/kotlin/test_apply_scope_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = "kotlin"
                .also { __check((it).toString(), "kotlin") }
                .let { it.reversed() }
                .run { toUpperCase() }
            __check((text).toString(), "NILTOK")
        }
