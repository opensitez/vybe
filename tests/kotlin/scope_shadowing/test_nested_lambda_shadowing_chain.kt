// vybe-test: kotlin/scope_shadowing/test_nested_lambda_shadowing_chain
// origin: languages/kotlin/tests/kotlin/test_scope_shadowing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val prefix = "A"
            val f = { prefix: String ->
                { prefix: Int -> "${'$'}{prefix}_${'$'}{prefix + 1}" }
            }
            val g = f("B")
            __check((g(3)).toString(), "B_4")
            __check((prefix).toString(), "A")
        }
