// vybe-test: kotlin/builtins/test_max_min_chain_behaviors
// origin: languages/kotlin/tests/kotlin/test_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val edge = max(min(-3, 8), min(2, -11))
            val span = min(max(9, 2), max(1, 4))
            __check((edge).toString(), "2")
            __check((span).toString(), "4")
        }
