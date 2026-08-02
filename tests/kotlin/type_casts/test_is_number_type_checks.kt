// vybe-test: kotlin/type_casts/test_is_number_type_checks
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val intValue: Any = 12
            val doubleValue: Any = 1.5
            __check((intValue is Int).toString(), "true")
            __check((doubleValue is Int).toString(), "false")
            __check((doubleValue is Double).toString(), "true")
        }
