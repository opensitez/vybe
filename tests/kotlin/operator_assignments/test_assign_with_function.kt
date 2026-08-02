// vybe-test: kotlin/operator_assignments/test_assign_with_function
// origin: languages/kotlin/tests/kotlin/test_operator_assignments.rs

fun add(a: Int): Int = a + 10
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
    var x = 1
    x += add(2)
    __check((x).toString(), "13")
}
