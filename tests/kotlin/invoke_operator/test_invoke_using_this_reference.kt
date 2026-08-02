// vybe-test: kotlin/invoke_operator/test_invoke_using_this_reference
// origin: languages/kotlin/tests/kotlin/test_invoke_operator.rs

class Counter {
            private var total = 0
            operator fun invoke(): Int {
                total += 1
                return total
            }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = Counter()
            __check((c()).toString(), "1")
            __check((c()).toString(), "2")
            __check((c()).toString(), "3")
        }
