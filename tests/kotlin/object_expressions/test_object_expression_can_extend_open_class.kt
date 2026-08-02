// vybe-test: kotlin/object_expressions/test_object_expression_can_extend_open_class
// origin: languages/kotlin/tests/kotlin/test_object_expressions.rs

open class Base {
            open fun label(): String = "base"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = object : Base() {
                override fun label(): String {
                    return super.label() + "-child"
                }
            }
            __check((value.label()).toString(), "base-child")
        }
