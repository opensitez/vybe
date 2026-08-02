// vybe-test: kotlin/operator_assignments/test_reassign_string_builder
// origin: languages/kotlin/tests/kotlin/test_operator_assignments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { var s = "x"
s += "y"
s += "z"
__check((s).toString(), "xyz") }
