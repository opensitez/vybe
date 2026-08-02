// vybe-test: kotlin/operators/test_is_and_not_is_checks
// origin: languages/kotlin/tests/kotlin/test_operators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Any? = "kotlin"
            __check((value is String).toString(), "true")
            __check((value is Int).toString(), "false")
            __check((value !is Int).toString(), "true")
            val nullValue: Any? = null
            __check((nullValue is String).toString(), "false")
            __check((nullValue !is String).toString(), "true")
        }
