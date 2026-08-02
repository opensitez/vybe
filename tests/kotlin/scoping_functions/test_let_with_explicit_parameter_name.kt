// vybe-test: kotlin/scoping_functions/test_let_with_explicit_parameter_name
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = "kotlin"
            val result = source.let { text -> text.uppercase() }
            __check((result).toString(), "KOTLIN")
        }
