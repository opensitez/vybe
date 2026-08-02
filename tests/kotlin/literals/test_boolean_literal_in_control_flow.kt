// vybe-test: kotlin/literals/test_boolean_literal_in_control_flow
// origin: languages/kotlin/tests/kotlin/test_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val yes = true
            val no = false
            val chosen = if (yes && !no) "go" else "stop"
            val result = if (yes && no) "bad" else if (!no) "ok" else "never"
            __check((chosen).toString(), "go")
            __check((result).toString(), "ok")
        }
