// vybe-test: kotlin/object_expressions/test_object_expression_with_arguments_like_member
// origin: languages/kotlin/tests/kotlin/test_object_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val adder = object {
                fun apply(base: Int, extra: Int): Int {
                    return base + extra
                }
            }
            __check((adder.apply(2, 5)).toString(), "7")
        }
