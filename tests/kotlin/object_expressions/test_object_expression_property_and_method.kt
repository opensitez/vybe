// vybe-test: kotlin/object_expressions/test_object_expression_property_and_method
// origin: languages/kotlin/tests/kotlin/test_object_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val obj = object { var value = 1
fun inc() { value += 1 } }
obj.inc()
obj.inc()
__check((obj.value).toString(), "3") }
