// vybe-test: kotlin/operator_assignments/test_nested_plus_assign
// origin: languages/kotlin/tests/kotlin/test_operator_assignments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { var a = 1
a += 2
a += 3
__check((a).toString(), "6") }
