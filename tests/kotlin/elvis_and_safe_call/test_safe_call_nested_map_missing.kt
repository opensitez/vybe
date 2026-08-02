// vybe-test: kotlin/elvis_and_safe_call/test_safe_call_nested_map_missing
// origin: languages/kotlin/tests/kotlin/test_elvis_and_safe_call.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        val x: Map<String, Map<String, Int>?>? = mapOf("a" to null)
        __check((x?.get("a")?.get("b") ?: -1).toString(), "-1")
    }
