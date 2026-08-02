// vybe-test: kotlin/scope/test_nested_block_rebinds_local_name
// origin: languages/kotlin/tests/kotlin/test_scope.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val token = "global"
            {
                val token = "inner"
                __check((token).toString(), "inner")
            }
            __check((token).toString(), "global")
        }
