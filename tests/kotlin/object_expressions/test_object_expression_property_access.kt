// vybe-test: kotlin/object_expressions/test_object_expression_property_access
// origin: languages/kotlin/tests/kotlin/test_object_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val record = object {
                val id = 10
                val label = "log"
            }
            __check((record.id).toString(), "10")
            __check((record.label).toString(), "log")
        }
