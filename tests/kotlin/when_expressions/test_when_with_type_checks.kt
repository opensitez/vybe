// vybe-test: kotlin/when_expressions/test_when_with_type_checks
// origin: languages/kotlin/tests/kotlin/test_when_expressions.rs

fun classify(value: Any): String {
            return when (value) {
                is Int -> "int"
                is String -> "string"
                else -> "other"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((classify(1)).toString(), "int")
            __check((classify("x")).toString(), "string")
            __check((classify(2.0)).toString(), "other")
        }
