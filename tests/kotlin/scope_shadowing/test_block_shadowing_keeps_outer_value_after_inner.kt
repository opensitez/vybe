// vybe-test: kotlin/scope_shadowing/test_block_shadowing_keeps_outer_value_after_inner
// origin: languages/kotlin/tests/kotlin/test_scope_shadowing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = "outer"
            val inside = run {
                val value = "inner"
                value
            }
            __check((inside).toString(), "inner")
            __check((value).toString(), "outer")
        }
