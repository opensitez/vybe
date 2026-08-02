// vybe-test: kotlin/try_finally/test_try_finally_without_return
// origin: languages/kotlin/tests/kotlin/test_try_finally.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        var out = "start"
        try {
            out = "try"
        } finally {
            out += "-fin"
        }
        __check((out).toString(), "try-fin")
    }
