// vybe-test: kotlin/elvis_and_safe_call/test_safe_call_in_when
// origin: languages/kotlin/tests/kotlin/test_elvis_and_safe_call.rs

fun main() {
        val x: String? = null
        when (x?.length ?: 0) {
            0 -> println("z")
            else -> println("n")
        }
    }

