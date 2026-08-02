// vybe-test: kotlin/repeat_statements/test_repeat_with_string_concat_condition
// origin: languages/kotlin/tests/kotlin/test_repeat_statements.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var out = ""
            repeat(4) { n ->
                out += if (n % 2 == 0) "E" else "O"
            }
            __check((out).toString(), "EOEO")
        }
