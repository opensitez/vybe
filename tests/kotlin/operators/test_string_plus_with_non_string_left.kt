// vybe-test: kotlin/operators/test_string_plus_with_non_string_left
// origin: languages/kotlin/tests/kotlin/test_operators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((2 + 3 + " apples").toString(), "5 apples")
            __check(("apples " + 2 + 3).toString(), "apples 23")
            __check(("calc " + (2 + 3)).toString(), "calc 5")
        }
