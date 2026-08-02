// vybe-test: kotlin/literals/test_boolean_literals
// origin: languages/kotlin/tests/kotlin/test_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = true
            val b = false
            __check((a && !b).toString(), "true")
            __check((a || b).toString(), "true")
            __check((!a && b).toString(), "false")
            __check((true == true).toString(), "true")
            __check((false == false).toString(), "true")
        }
