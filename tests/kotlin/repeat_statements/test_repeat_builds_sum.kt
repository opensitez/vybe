// vybe-test: kotlin/repeat_statements/test_repeat_builds_sum
// origin: languages/kotlin/tests/kotlin/test_repeat_statements.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var sum = 0
            repeat(5) { i ->
                sum += i
            }
            __check((sum).toString(), "10")
        }
