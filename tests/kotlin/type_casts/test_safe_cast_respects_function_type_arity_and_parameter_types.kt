// vybe-test: kotlin/type_casts/test_safe_cast_respects_function_type_arity_and_parameter_types
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val handler: Any = { value: Int -> value.toString() }
            val unary = handler as? (Int) -> String
            val binary = handler as? (Int, Int) -> String
            __check((unary != null).toString(), "true")
            __check((binary == null).toString(), "true")
        }
