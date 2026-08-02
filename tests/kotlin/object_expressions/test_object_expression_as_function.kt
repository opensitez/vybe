// vybe-test: kotlin/object_expressions/test_object_expression_as_function
// origin: languages/kotlin/tests/kotlin/test_object_expressions.rs

fun make(): Int { val f = object { fun call(v: Int) = v * 3 }
return f.call(2) }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((make()).toString(), "6") }
