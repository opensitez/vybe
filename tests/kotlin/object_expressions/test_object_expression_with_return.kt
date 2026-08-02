// vybe-test: kotlin/object_expressions/test_object_expression_with_return
// origin: languages/kotlin/tests/kotlin/test_object_expressions.rs

fun makeOutput() = object { fun out() = "pong" }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val o = makeOutput()
__check((o.out()).toString(), "pong") }
