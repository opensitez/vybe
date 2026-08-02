// vybe-test: kotlin/scope_shadowing/test_let_shadowing_chain
// origin: languages/kotlin/tests/kotlin/test_scope_shadowing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var value = "outer"
            val result = value.let { value ->
                val value = value + ":inner"
                value
            }
            __check((result).toString(), "outer:inner")
            __check((value).toString(), "outer")
        }
