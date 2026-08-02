// vybe-test: kotlin/control_flow/test_when_as_expression_in_assignment
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val score = 77
            val grade = when (score) {
                in 90..100 -> "A"
                in 80..89 -> "B"
                in 70..79 -> "C"
                else -> "F"
            }
            __check((grade).toString(), "C")
        }
