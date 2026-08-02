// vybe-test: kotlin/object_expressions/test_object_expression_nested_creation
// origin: languages/kotlin/tests/kotlin/test_object_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val builder = object {
                fun create(): String {
                    return "x"
                }
            }

            val name = builder.create()
            val wrapper = object {
                fun wrap(value: String): String {
                    return value + value
                }
            }

            __check((wrapper.wrap(name)).toString(), "xx")
        }
