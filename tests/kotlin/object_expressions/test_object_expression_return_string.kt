// vybe-test: kotlin/object_expressions/test_object_expression_return_string
// origin: languages/kotlin/tests/kotlin/test_object_expressions.rs

fun makeLabel(): String { val obj = object { fun text() = "ok" }
return obj.text() }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((makeLabel()).toString(), "ok") }
