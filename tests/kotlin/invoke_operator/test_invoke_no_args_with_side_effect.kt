// vybe-test: kotlin/invoke_operator/test_invoke_no_args_with_side_effect
// origin: languages/kotlin/tests/kotlin/test_invoke_operator.rs

var count = 0
        class Notifier {
            operator fun invoke() {
                count += 1
            }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val n = Notifier()
            n()
            n()
            __check((count).toString(), "2")
        }
