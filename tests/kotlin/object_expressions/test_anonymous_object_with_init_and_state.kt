// vybe-test: kotlin/object_expressions/test_anonymous_object_with_init_and_state
// origin: languages/kotlin/tests/kotlin/test_object_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val stateful = object {
                var value = 0
                fun inc() {
                    value += 1
                }
                init {
                    value = 5
                }
            }

            stateful.inc()
            __check((stateful.value).toString(), "6")
        }
