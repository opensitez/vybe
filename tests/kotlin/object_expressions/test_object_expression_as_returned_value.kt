// vybe-test: kotlin/object_expressions/test_object_expression_as_returned_value
// origin: languages/kotlin/tests/kotlin/test_object_expressions.rs

interface Producer {
            fun value(): Int
        }

        fun makeProducer(start: Int): Producer {
            return object : Producer {
                override fun value(): Int {
                    return start
                }
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val p = makeProducer(9)
            __check((p.value()).toString(), "9")
        }
