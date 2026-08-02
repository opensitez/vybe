// vybe-test: kotlin/object_expressions/test_object_expression_as_nested_return_value
// origin: languages/kotlin/tests/kotlin/test_object_expressions.rs

interface Calculator {
            fun add(value: Int): Int
        }

        fun wrap(base: Int): Calculator {
            return object : Calculator {
                override fun add(value: Int): Int {
                    return base + value
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
            val calc = wrap(4)
            __check((calc.add(3)).toString(), "7")
            __check((calc.add(1)).toString(), "5")
        }
