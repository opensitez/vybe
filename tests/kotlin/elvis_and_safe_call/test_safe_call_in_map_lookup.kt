// vybe-test: kotlin/elvis_and_safe_call/test_safe_call_in_map_lookup
// origin: languages/kotlin/tests/kotlin/test_elvis_and_safe_call.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        val m: Map<String, String>? = null
        __check((m?.get("a") ?: "na").toString(), "na")
    }
