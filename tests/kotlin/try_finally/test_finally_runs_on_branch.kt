// vybe-test: kotlin/try_finally/test_finally_runs_on_branch
// origin: languages/kotlin/tests/kotlin/test_try_finally.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        val v = if (true) {
            try { "yes" } finally { __check(("fin").toString(), "fin") }
        } else { "no" }
        __check((v).toString(), "yes")
    }
