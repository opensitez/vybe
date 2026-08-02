// vybe-test: kotlin/scoping_functions/test_let_on_nullable_returns_none_when_null
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Int? = null
            val mapped = value?.let { it + 1 }
            __check((mapped == null).toString(), "true")
        }
