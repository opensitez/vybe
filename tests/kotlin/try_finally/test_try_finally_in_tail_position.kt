// vybe-test: kotlin/try_finally/test_try_finally_in_tail_position
// origin: languages/kotlin/tests/kotlin/test_try_finally.rs

fun f(v: Int): Int = try { v + 1 } finally { __check(("fin").toString(), "fin") }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((f(3)).toString(), "4") }
