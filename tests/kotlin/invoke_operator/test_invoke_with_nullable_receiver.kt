// vybe-test: kotlin/invoke_operator/test_invoke_with_nullable_receiver
// origin: languages/kotlin/tests/kotlin/test_invoke_operator.rs

class Greeter {
            operator fun invoke(name: String?): String = name ?: "guest"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val g = Greeter()
            __check((g(null)).toString(), "guest")
        }
