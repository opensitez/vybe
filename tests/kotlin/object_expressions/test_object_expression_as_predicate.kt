// vybe-test: kotlin/object_expressions/test_object_expression_as_predicate
// origin: languages/kotlin/tests/kotlin/test_object_expressions.rs

interface Check { fun ok(v: Int): Boolean }
fun runCheck(c: Check): Boolean = c.ok(5)
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((runCheck(object : Check { override fun ok(v: Int) = v > 3 })).toString(), "true") }
