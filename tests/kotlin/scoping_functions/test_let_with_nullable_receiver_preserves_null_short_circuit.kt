// vybe-test: kotlin/scoping_functions/test_let_with_nullable_receiver_preserves_null_short_circuit
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: String? = null
            val projected = value?.let {
                "inside"
            }
            __check((projected == null).toString(), "true")
            __check((value).toString(), "null")
        }
