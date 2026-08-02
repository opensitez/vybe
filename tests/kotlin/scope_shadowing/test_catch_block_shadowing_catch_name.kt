// vybe-test: kotlin/scope_shadowing/test_catch_block_shadowing_catch_name
// origin: languages/kotlin/tests/kotlin/test_scope_shadowing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val e = "outer"
            try {
                throw IllegalStateException("boom")
            } catch (e: Exception) {
                __check((e.message).toString(), "boom")
            }
            __check((e).toString(), "outer")
        }
