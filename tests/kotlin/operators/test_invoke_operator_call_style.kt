// vybe-test: kotlin/operators/test_invoke_operator_call_style
// origin: languages/kotlin/tests/kotlin/test_operators.rs

class Transformer {
            operator fun invoke(value: Int): Int {
                return value * value
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val transform = Transformer()
            __check((transform(4)).toString(), "16")
        }
