// vybe-test: kotlin/elvis_and_safe_call/test_notnull_cast_and_safe
// origin: languages/kotlin/tests/kotlin/test_elvis_and_safe_call.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        val x: Any? = "ok"
        __check(((x as? String)?.length ?: 0).toString(), "2")
    }
