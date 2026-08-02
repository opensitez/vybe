// vybe-test: kotlin/invoke_operator/test_invoke_no_args
// origin: languages/kotlin/tests/kotlin/test_invoke_operator.rs

class Counter {
            operator fun invoke(): Int = 1
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Counter()()).toString(), "1")
        }
