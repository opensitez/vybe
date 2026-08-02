// vybe-test: kotlin/object_expressions/test_object_expression_with_state
// origin: languages/kotlin/tests/kotlin/test_object_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val counter = object {
                var value = 0
                fun inc() {
                    value += 1
                }
                fun reset() {
                    value = 0
                }
            }

            counter.inc()
            counter.inc()
            counter.reset()
            counter.inc()
            __check((counter.value).toString(), "1")
        }
