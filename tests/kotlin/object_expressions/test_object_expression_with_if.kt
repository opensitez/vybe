// vybe-test: kotlin/object_expressions/test_object_expression_with_if
// origin: languages/kotlin/tests/kotlin/test_object_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val obj = object {
                fun flag(x: Int): String {
                    if (x > 0) {
                        return "yes"
                    }
                    return "no"
                }
            }
            __check((obj.flag(1)).toString(), "yes")
            __check((obj.flag(-1)).toString(), "no")
        }
