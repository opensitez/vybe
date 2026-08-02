// vybe-test: kotlin/type_casts/test_is_check_for_number_interface_vs_concrete_numeric_type
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val intValue: Any = 12
            val longValue: Any = 12L
            val doubleValue: Any = 12.0

            __check((intValue is Number).toString(), "true")
            __check((longValue is Int).toString(), "false")
            __check((doubleValue is Long).toString(), "false")
            __check((longValue is Number).toString(), "true")
            __check((intValue as? Int != null).toString(), "true")
            __check((longValue as? Int == null).toString(), "true")
        }
