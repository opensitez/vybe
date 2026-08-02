// vybe-test: kotlin/object_expressions/test_object_expression_with_args
// origin: languages/kotlin/tests/kotlin/test_object_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val sum = object { fun add(x: Int, y: Int) = x + y }
__check((sum.add(4, 5)).toString(), "9") }
