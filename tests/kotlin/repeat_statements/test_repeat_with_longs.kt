// vybe-test: kotlin/repeat_statements/test_repeat_with_longs
// origin: languages/kotlin/tests/kotlin/test_repeat_statements.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var total = 0L
            repeat(4) { i ->
                total += i.toLong()
            }
            __check((total).toString(), "6")
        }
