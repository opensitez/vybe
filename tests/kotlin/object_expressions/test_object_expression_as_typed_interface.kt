// vybe-test: kotlin/object_expressions/test_object_expression_as_typed_interface
// origin: languages/kotlin/tests/kotlin/test_object_expressions.rs

interface Worker {
            fun work(): String
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val w: Worker = object : Worker {
                override fun work(): String {
                    return "done"
                }
            }
            __check((w.work()).toString(), "done")
        }
