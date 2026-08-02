// vybe-test: kotlin/object_expressions/test_object_expression_with_custom_getter_and_setter
// origin: languages/kotlin/tests/kotlin/test_object_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val obj = object {
                var counter = 0
                val doubled: Int
                    get() = counter * 2
            }
            obj.counter = 5
            __check((obj.doubled).toString(), "10")
        }
