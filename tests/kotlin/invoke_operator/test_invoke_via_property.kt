// vybe-test: kotlin/invoke_operator/test_invoke_via_property
// origin: languages/kotlin/tests/kotlin/test_invoke_operator.rs

class Builder {
            val call = { n: Int -> n * 2 }
            operator fun invoke(v: Int): Int = call(v)
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val b = Builder()
            __check((b(4)).toString(), "8")
        }
