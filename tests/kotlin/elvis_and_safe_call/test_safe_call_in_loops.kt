// vybe-test: kotlin/elvis_and_safe_call/test_safe_call_in_loops
// origin: languages/kotlin/tests/kotlin/test_elvis_and_safe_call.rs

fun main() {
        val xs: List<String?> = listOf(null, "a", null, "bb")
        var c = 0
        for (s in xs) { c += s?.length ?: 0 }
        println(c)
    }

