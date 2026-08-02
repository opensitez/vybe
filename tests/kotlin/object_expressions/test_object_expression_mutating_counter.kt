// vybe-test: kotlin/object_expressions/test_object_expression_mutating_counter
// origin: languages/kotlin/tests/kotlin/test_object_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val counter = object {
                var value = 1
                fun inc() {
                    value *= 2
                }
            }

            counter.inc()
            counter.inc()
            __check((counter.value).toString(), "4")
        }
