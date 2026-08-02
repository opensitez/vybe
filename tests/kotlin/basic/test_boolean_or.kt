// vybe-test: kotlin/basic/test_boolean_or
// origin: languages/kotlin/tests/kotlin/test_basic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((true || false).toString(), "true")
            __check((false || false).toString(), "false")
        }
