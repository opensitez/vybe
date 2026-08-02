// vybe-test: kotlin/object_expressions/test_object_expression_chain
// origin: languages/kotlin/tests/kotlin/test_object_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val a = object { fun first() = 2 }
val b = object { fun second() = 3 }
__check((a.first() + b.second()).toString(), "5") }
