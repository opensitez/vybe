// vybe-test: kotlin/scoping_functions/test_scoped_chain_mix_with_take_if_for_filtering
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base = "value".takeIf { it.length > 2 } ?: "none"
            val result = base.let { it + "-ok" }
            __check((result).toString(), "value-ok")
        }
