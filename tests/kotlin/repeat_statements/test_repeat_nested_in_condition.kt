// vybe-test: kotlin/repeat_statements/test_repeat_nested_in_condition
// origin: languages/kotlin/tests/kotlin/test_repeat_statements.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var out = 0
            repeat(4) { outer ->
                if (outer == 2) {
                    repeat(2) {
                        out += outer
                    }
                }
            }
            __check((out).toString(), "4")
        }
