// vybe-test: kotlin/object_expressions/test_object_expression_stored_and_reused
// origin: languages/kotlin/tests/kotlin/test_object_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val provider = object {
                fun value(): Int {
                    return 3
                }
            }

            val first = provider.value()
            val second = provider.value()
            __check((first + second).toString(), "6")
        }
