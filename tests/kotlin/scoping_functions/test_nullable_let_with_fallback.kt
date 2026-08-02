// vybe-test: kotlin/scoping_functions/test_nullable_let_with_fallback
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Int? = null
            val mapped = value?.let { it + 1 } ?: -1
            __check((mapped).toString(), "-1")
        }
