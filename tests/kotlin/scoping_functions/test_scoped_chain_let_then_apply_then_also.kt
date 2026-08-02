// vybe-test: kotlin/scoping_functions/test_scoped_chain_let_then_apply_then_also
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val log = mutableListOf<String>()
            val result = "ok".let { it.uppercase() }
                .also { log.add("a") }
                .let { it + "-done" }
                .also { log.add("b") }
            __check((result).toString(), "OK-done")
            __check((log.joinToString("-")).toString(), "a-b")
        }
