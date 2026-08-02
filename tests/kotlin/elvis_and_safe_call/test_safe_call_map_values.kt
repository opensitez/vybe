// vybe-test: kotlin/elvis_and_safe_call/test_safe_call_map_values
// origin: languages/kotlin/tests/kotlin/test_elvis_and_safe_call.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        val x: Map<String, Int>? = mapOf("a" to 2)
        __check((x?.get("a") ?: -1).toString(), "2")
    }
