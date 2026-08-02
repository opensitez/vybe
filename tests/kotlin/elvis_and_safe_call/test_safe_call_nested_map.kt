// vybe-test: kotlin/elvis_and_safe_call/test_safe_call_nested_map
// origin: languages/kotlin/tests/kotlin/test_elvis_and_safe_call.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        val x: Map<String, Map<String, Int>?>? = mapOf("a" to mapOf("b" to 3))
        __check((x?.get("a")?.get("b") ?: -1).toString(), "3")
    }
