// vybe-test: kotlin/object_expressions/test_object_expression_two_methods
// origin: languages/kotlin/tests/kotlin/test_object_expressions.rs

interface MathOp { fun a(): Int
fun b(): Int }
fun makeOps() = object : MathOp { override fun a() = 2
override fun b() = 3 }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val op = makeOps()
__check((op.a() + op.b()).toString(), "5") }
