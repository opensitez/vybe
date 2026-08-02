// vybe-test: kotlin/object_expressions/test_object_expression_double_field
// origin: languages/kotlin/tests/kotlin/test_object_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val pair = object { var left = 1
var right = 2 }
pair.left += 3
__check((pair.left + pair.right).toString(), "6") }
