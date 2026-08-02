// vybe-test: kotlin/exceptions/test_exception_check_like_guard
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

fun validate(ok: Boolean): Int {
    return if (ok) 1 else 0
}

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
    try {
        if (validate(false) == 0) {
            throw Exception("bad")
        }
    } catch (e: Exception) {
        __check(("bad").toString(), "bad")
    }
}
