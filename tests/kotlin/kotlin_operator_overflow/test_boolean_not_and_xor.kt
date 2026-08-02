// vybe-test: kotlin/kotlin_operator_overflow/test_boolean_not_and_xor
// origin: languages/kotlin/tests/kotlin/test_kotlin_operator_overflow.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = true
            val b = false
            __check((!a).toString(), "false")
            __check((a xor b).toString(), "true")
        }
