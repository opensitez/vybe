// vybe-test: kotlin/object_expressions/test_object_expression_with_state_reset
// origin: languages/kotlin/tests/kotlin/test_object_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val obj = object { var value = 0
fun add(v: Int) { value += v }
fun reset() { value = 0 } }
obj.add(5)
obj.reset()
__check((obj.value).toString(), "0") }
