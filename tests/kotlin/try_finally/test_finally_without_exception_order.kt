// vybe-test: kotlin/try_finally/test_finally_without_exception_order
// origin: languages/kotlin/tests/kotlin/test_try_finally.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        val out = run {
            var x = ""
            try {
                x += "t"
            } finally {
                x += "f"
            }
            x
        }
        __check((out).toString(), "tf")
    }
