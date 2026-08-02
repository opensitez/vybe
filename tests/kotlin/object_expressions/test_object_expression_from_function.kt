// vybe-test: kotlin/object_expressions/test_object_expression_from_function
// origin: languages/kotlin/tests/kotlin/test_object_expressions.rs

fun create(): Int { val worker = object { fun run(v: Int) = v + 1 }
return worker.run(4) }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((create()).toString(), "5") }
