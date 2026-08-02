// vybe-test: kotlin/scope/test_block_reused_name_after_scope
// origin: languages/kotlin/tests/kotlin/test_scope.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var name = "outer"
            {
                val name = "inner"
                __check((name).toString(), "inner")
            }
            name = "next"
            __check((name).toString(), "next")
        }
