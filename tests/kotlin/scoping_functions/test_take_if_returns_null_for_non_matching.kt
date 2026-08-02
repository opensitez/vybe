// vybe-test: kotlin/scoping_functions/test_take_if_returns_null_for_non_matching
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((4.takeIf { it > 10 } == null).toString(), "true")
            __check((4.takeIf { it < 10 } == null).toString(), "false")
        }
