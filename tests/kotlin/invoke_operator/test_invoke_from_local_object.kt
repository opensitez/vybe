// vybe-test: kotlin/invoke_operator/test_invoke_from_local_object
// origin: languages/kotlin/tests/kotlin/test_invoke_operator.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val f = object {
                operator fun invoke(v: Int): Int = v * v
            }
            __check((f(6)).toString(), "36")
        }
