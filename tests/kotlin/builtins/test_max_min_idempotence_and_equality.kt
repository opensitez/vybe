// vybe-test: kotlin/builtins/test_max_min_idempotence_and_equality
// origin: languages/kotlin/tests/kotlin/test_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((max(7, 7)).toString(), "7")
            __check((min(7, 7)).toString(), "7")
            __check((max(-3, -3)).toString(), "-3")
            __check((min(-3, -3)).toString(), "-3")
        }
