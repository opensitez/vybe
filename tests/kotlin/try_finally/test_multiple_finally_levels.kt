// vybe-test: kotlin/try_finally/test_multiple_finally_levels
// origin: languages/kotlin/tests/kotlin/test_try_finally.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        try {
            __check(("outer-try").toString(), "outer-try")
            try {
                __check(("inner-try").toString(), "inner-try")
            } finally {
                __check(("inner-f").toString(), "inner-f")
            }
        } finally {
            __check(("outer-f").toString(), "outer-f")
        }
    }
