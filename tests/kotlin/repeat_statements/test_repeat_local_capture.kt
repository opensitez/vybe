// vybe-test: kotlin/repeat_statements/test_repeat_local_capture
// origin: languages/kotlin/tests/kotlin/test_repeat_statements.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var label = ""
            repeat(4) { i ->
                if (i > 1) label += i.toString()
            }
            __check((label).toString(), "23")
        }
