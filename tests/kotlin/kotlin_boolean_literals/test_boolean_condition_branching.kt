// vybe-test: kotlin/kotlin_boolean_literals/test_boolean_condition_branching
// origin: languages/kotlin/tests/kotlin/test_kotlin_boolean_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x = if (true) {
                1
            } else {
                2
            }
            __check((x).toString(), "1")
        }
