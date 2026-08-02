// vybe-test: kotlin/primitive_array_apis/test_boolean_array_any_all_none
// origin: languages/kotlin/tests/kotlin/test_primitive_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = booleanArrayOf(true, false, true)
            __check((values.any()).toString(), "true")
            __check((values.all { it }).toString(), "false")
            __check((values.none { it == null }).toString(), "true")
        }
