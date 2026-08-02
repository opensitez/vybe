// vybe-test: kotlin/basic/test_boolean_not
// origin: languages/kotlin/tests/kotlin/test_basic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((!true).toString(), "false")
            __check((!false).toString(), "true")
        }
