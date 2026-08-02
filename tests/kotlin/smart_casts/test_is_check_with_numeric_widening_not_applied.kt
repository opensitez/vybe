// vybe-test: kotlin/smart_casts/test_is_check_with_numeric_widening_not_applied
// origin: languages/kotlin/tests/kotlin/test_smart_casts.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Any = 7
            __check((value is Int).toString(), "true")
            __check((value is Long).toString(), "false")
        }
