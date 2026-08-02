// vybe-test: kotlin/try_finally/test_try_finally_in_function
// origin: languages/kotlin/tests/kotlin/test_try_finally.rs

fun run(): Int {
        var x = 0
        try { x = 1 } finally { x = 2 }
        return x
    }
    fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((run()).toString(), "2") }
