// vybe-test: kotlin/object_expressions/test_object_expression_with_interface
// origin: languages/kotlin/tests/kotlin/test_object_expressions.rs

interface Callback {
            fun onComplete()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val cb = object : Callback {
                override fun onComplete() {
                    __check(("callback finished").toString(), "callback finished")
                }
            }
            cb.onComplete()
        }
