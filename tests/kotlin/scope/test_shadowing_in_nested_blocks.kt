// vybe-test: kotlin/scope/test_shadowing_in_nested_blocks
// origin: languages/kotlin/tests/kotlin/test_scope.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val mode = "outer"
            __check((mode).toString(), "outer")
            {
                val mode = "inner"
                __check((mode).toString(), "inner")
            }
            __check((mode).toString(), "outer")
        }
