// vybe-test: kotlin/scope/test_function_literal_scope_with_block_shadowing
// origin: languages/kotlin/tests/kotlin/test_scope.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 1
            val add = { input: Int ->
                val result = input + 1
                result
            }
            __check((add(3)).toString(), "4")
            __check((value).toString(), "1")
        }
