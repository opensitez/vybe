// vybe-test: kotlin/operator_assignments/test_assign_inside_function
// origin: languages/kotlin/tests/kotlin/test_operator_assignments.rs

fun bump(v: Int): Int { var x = v
x += 1
return x }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((bump(8)).toString(), "9") }
