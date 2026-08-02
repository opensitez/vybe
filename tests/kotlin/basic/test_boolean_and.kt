// vybe-test: kotlin/basic/test_boolean_and
// origin: languages/kotlin/tests/kotlin/test_basic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((true && true).toString(), "true")
            __check((true && false).toString(), "false")
        }
