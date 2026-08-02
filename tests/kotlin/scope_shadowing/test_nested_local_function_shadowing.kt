// vybe-test: kotlin/scope_shadowing/test_nested_local_function_shadowing
// origin: languages/kotlin/tests/kotlin/test_scope_shadowing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            fun outer(): String {
                val value = "outer"
                fun inner(): String {
                    val value = "inner"
                    return value
                }
                return "${'$'}{inner()}|${'$'}value"
            }
            __check((outer()).toString(), "inner|outer")
        }
