// vybe-test: kotlin/object_expressions/test_object_expression_local_type
// origin: languages/kotlin/tests/kotlin/test_object_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val result = object { val value = 1 }
__check((result.value + 4).toString(), "5") }
