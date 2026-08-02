// vybe-test: kotlin/when_expressions/test_when_with_non_exhaustive_else_on_any_subject
// origin: languages/kotlin/tests/kotlin/test_when_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((when ("x") {
                "a" -> 1
                "b" -> 2
                else -> 3
            }).toString(), "3")
        }
