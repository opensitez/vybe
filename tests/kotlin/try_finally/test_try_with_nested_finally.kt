// vybe-test: kotlin/try_finally/test_try_with_nested_finally
// origin: languages/kotlin/tests/kotlin/test_try_finally.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        try {
            try { __check(("inner").toString(), "inner") }
            finally { __check(("inner-finally").toString(), "inner-finally") }
        } finally { __check(("outer-finally").toString(), "outer-finally") }
    }
