// vybe-test: kotlin/numeric_types/test_compound_mod_assign
// origin: languages/kotlin/tests/kotlin/test_numeric_types.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var value = 23
            value %= 5
            __check((value).toString(), "3")
            value %= 2
            __check((value).toString(), "1")
        }
