// vybe-test: kotlin/repeat_statements/test_repeat_in_class_init
// origin: languages/kotlin/tests/kotlin/test_repeat_statements.rs

class Counter {
            var total = 0
            init {
                repeat(3) { total += 1 }
            }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Counter().total).toString(), "3")
        }
