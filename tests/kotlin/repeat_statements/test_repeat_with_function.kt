// vybe-test: kotlin/repeat_statements/test_repeat_with_function
// origin: languages/kotlin/tests/kotlin/test_repeat_statements.rs

fun addOne(x: Int) = x + 1
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var total = 0
            repeat(4) {
                total = addOne(total)
            }
            __check((total).toString(), "4")
        }
