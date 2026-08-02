// vybe-test: kotlin/block_expressions/test_block_assign_to_property
// origin: languages/kotlin/tests/kotlin/test_block_expressions.rs

class X { var value = 1 }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val x = X()
x.value = run { val a = x.value
a + 4 }
__check((x.value).toString(), "5") }
