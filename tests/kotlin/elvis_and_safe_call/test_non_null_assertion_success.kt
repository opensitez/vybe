// vybe-test: kotlin/elvis_and_safe_call/test_non_null_assertion_success
// origin: languages/kotlin/tests/kotlin/test_elvis_and_safe_call.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val x: String? = "ok"
__check((x!!).toString(), "ok")
}
