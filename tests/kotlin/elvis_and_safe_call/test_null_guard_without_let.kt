// vybe-test: kotlin/elvis_and_safe_call/test_null_guard_without_let
// origin: languages/kotlin/tests/kotlin/test_elvis_and_safe_call.rs

fun main() {
        val x: String? = null
        x?.let { println("v" + it) } ?: println("missing")
    }

