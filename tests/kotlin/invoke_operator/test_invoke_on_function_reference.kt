// vybe-test: kotlin/invoke_operator/test_invoke_on_function_reference
// origin: languages/kotlin/tests/kotlin/test_invoke_operator.rs

class Adder {
            operator fun invoke(x: Int): Int = x + 1
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val fn: (Int) -> Int = Adder()
            __check((fn(2)).toString(), "3")
        }
