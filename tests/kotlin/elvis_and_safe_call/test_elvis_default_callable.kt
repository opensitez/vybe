// vybe-test: kotlin/elvis_and_safe_call/test_elvis_default_callable
// origin: languages/kotlin/tests/kotlin/test_elvis_and_safe_call.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        fun choose(v: String?): String = v ?: run { __check(("fallback").toString(), "fallback")
"x" }
        __check((choose(null)).toString(), "x")
    }
