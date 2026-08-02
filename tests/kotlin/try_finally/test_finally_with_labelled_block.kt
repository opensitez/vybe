// vybe-test: kotlin/try_finally/test_finally_with_labelled_block
// origin: languages/kotlin/tests/kotlin/test_try_finally.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        try {
            __check(("start").toString(), "start")
        } finally { __check(("finally").toString(), "finally") }
    }
