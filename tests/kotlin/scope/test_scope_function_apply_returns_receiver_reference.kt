// vybe-test: kotlin/scope/test_scope_function_apply_returns_receiver_reference
// origin: languages/kotlin/tests/kotlin/test_scope.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = StringBuilder()
                .apply {
                    append("k")
                    append("otlin")
                }
            __check((text.toString()).toString(), "kotlin")
            __check((text === StringBuilder("kotlin")).toString(), "false")
        }
