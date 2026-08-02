// vybe-test: kotlin/apply_scope_functions/test_scoping_function_on_nullable
// origin: languages/kotlin/tests/kotlin/test_apply_scope_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: String? = "abc"
            val a = value?.let { it + "d" }
            val b = value?.let { null }
            __check((a).toString(), "abcd")
            __check((b).toString(), "null")
        }
