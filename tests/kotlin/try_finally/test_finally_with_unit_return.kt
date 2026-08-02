// vybe-test: kotlin/try_finally/test_finally_with_unit_return
// origin: languages/kotlin/tests/kotlin/test_try_finally.rs

fun run(): Unit {
        try {
            __check(("u").toString(), "u")
        } finally {
            __check(("f").toString(), "f")
        }
    }
    fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { run() }
