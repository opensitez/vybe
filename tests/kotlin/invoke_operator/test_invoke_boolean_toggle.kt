// vybe-test: kotlin/invoke_operator/test_invoke_boolean_toggle
// origin: languages/kotlin/tests/kotlin/test_invoke_operator.rs

class Toggle {
            var state = false
            operator fun invoke(): Boolean {
                state = !state
                return state
            }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val t = Toggle()
            __check((t()).toString(), "true")
            __check((t()).toString(), "false")
        }
