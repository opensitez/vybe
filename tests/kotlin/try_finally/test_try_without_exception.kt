// vybe-test: kotlin/try_finally/test_try_without_exception
// origin: languages/kotlin/tests/kotlin/test_try_finally.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { try { __check(("ok").toString(), "ok") } finally { __check(("fin").toString(), "fin") } }
