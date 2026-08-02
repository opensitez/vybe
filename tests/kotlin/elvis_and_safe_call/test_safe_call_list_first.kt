// vybe-test: kotlin/elvis_and_safe_call/test_safe_call_list_first
// origin: languages/kotlin/tests/kotlin/test_elvis_and_safe_call.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        val xs: List<String>? = listOf("a", "b")
        __check((xs?.firstOrNull() ?: "none").toString(), "a")
    }
