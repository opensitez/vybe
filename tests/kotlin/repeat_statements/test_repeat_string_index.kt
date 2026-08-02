// vybe-test: kotlin/repeat_statements/test_repeat_string_index
// origin: languages/kotlin/tests/kotlin/test_repeat_statements.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var out = ""
            repeat(3) { index ->
                out += index.toString()
            }
            __check((out).toString(), "012")
        }
