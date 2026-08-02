// vybe-test: kotlin/object_expressions/test_object_expression_boolean_case
// origin: languages/kotlin/tests/kotlin/test_object_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val o = object { fun check(v: Int) = v % 2 == 0 }
__check((o.check(2)).toString(), "true")
__check((o.check(7)).toString(), "false") }
