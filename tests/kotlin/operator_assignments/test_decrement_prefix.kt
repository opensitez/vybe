// vybe-test: kotlin/operator_assignments/test_decrement_prefix
// origin: languages/kotlin/tests/kotlin/test_operator_assignments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { var a = 8
val b = --a
__check((a).toString(), "7")
__check((b).toString(), "7") }
