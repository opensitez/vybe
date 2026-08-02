// vybe-test: kotlin/control_flow/test_when_with_is_type_checks
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Any = "kotlin"
            val category = when (value) {
                is Int -> "int"
                is String -> "string"
                is Boolean -> "bool"
                else -> "other"
            }
            __check((category).toString(), "string")
        }
