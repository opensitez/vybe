// vybe-test: kotlin/elvis_and_safe_call/test_elvis_return
// origin: languages/kotlin/tests/kotlin/test_elvis_and_safe_call.rs

fun title(v: String?): String {
        return v ?: return "none"
    }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((title(null)).toString(), "none")
__check((title("x")).toString(), "x") }
