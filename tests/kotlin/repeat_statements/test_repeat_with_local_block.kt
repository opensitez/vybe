// vybe-test: kotlin/repeat_statements/test_repeat_with_local_block
// origin: languages/kotlin/tests/kotlin/test_repeat_statements.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var out = ""
            repeat(3) {
                out += "x"
            }
            __check((out).toString(), "xxx")
        }
