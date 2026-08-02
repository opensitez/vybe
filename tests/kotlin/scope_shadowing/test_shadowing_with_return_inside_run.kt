// vybe-test: kotlin/scope_shadowing/test_shadowing_with_return_inside_run
// origin: languages/kotlin/tests/kotlin/test_scope_shadowing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val marker = "outer"
            val result = run {
                val marker = "inner"
                marker
            }
            __check((result).toString(), "inner")
            __check((marker).toString(), "outer")
        }
