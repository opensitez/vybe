// vybe-test: kotlin/preconditions/test_check_true
// origin: languages/kotlin/tests/kotlin/test_preconditions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            check(1 + 1 == 2)
            __check(("pass").toString(), "pass")
        }
