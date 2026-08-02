// vybe-test: kotlin/elvis_and_safe_call/test_non_null_assertion_failure
// origin: languages/kotlin/tests/kotlin/test_elvis_and_safe_call.rs

fun main() { try { val x: String? = null
println(x!!) } catch (e: Exception) { println("err") } }

