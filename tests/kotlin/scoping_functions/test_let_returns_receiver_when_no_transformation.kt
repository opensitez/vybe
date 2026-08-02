// vybe-test: kotlin/scoping_functions/test_let_returns_receiver_when_no_transformation
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = "kotlin"
            val result = value.let { it }
            __check((result).toString(), "kotlin")
            __check((result === value).toString(), "true")
        }
