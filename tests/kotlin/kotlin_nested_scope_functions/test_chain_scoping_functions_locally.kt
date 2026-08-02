// vybe-test: kotlin/kotlin_nested_scope_functions/test_chain_scoping_functions_locally
// origin: languages/kotlin/tests/kotlin/test_kotlin_nested_scope_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = listOf(1, 2, 3)
                .map { it * 2 }
                .let { numbers -> numbers.filter { it > 2 } }
                .also { __check((it.size).toString(), "3") }
                .sum()
            __check((out).toString(), "12")
        }
