// vybe-test: kotlin/try_finally/test_mutation_in_finally_observable
// origin: languages/kotlin/tests/kotlin/test_try_finally.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        var x = 0
        try {
            __check(("start").toString(), "start")
        } finally {
            x = 9
            __check((x).toString(), "9")
        }
        __check((x).toString(), "9")
    }
