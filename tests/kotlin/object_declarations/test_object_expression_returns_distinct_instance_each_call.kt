// vybe-test: kotlin/object_declarations/test_object_expression_returns_distinct_instance_each_call
// origin: languages/kotlin/tests/kotlin/test_object_declarations.rs

fun builder(): Any {
            return object {
                val value = 1
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val first = builder()
            val second = builder()
            __check((first === second).toString(), "false")
        }
