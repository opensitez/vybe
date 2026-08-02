// vybe-test: kotlin/scope_shadowing/test_shadowed_variable_in_map_lambda
// origin: languages/kotlin/tests/kotlin/test_scope_shadowing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 1
            val result = listOf(1, 2, 3).map { value -> value * 2 }
            __check((result.joinToString(",")).toString(), "2,4,6")
            __check((value).toString(), "1")
        }
