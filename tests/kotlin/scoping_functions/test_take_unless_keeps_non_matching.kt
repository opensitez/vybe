// vybe-test: kotlin/scoping_functions/test_take_unless_keeps_non_matching
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 7.takeUnless { it == 0 }
            __check((value).toString(), "7")
        }
