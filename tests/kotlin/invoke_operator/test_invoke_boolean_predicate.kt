// vybe-test: kotlin/invoke_operator/test_invoke_boolean_predicate
// origin: languages/kotlin/tests/kotlin/test_invoke_operator.rs

class Matcher {
            operator fun invoke(v: Int): Boolean = v % 2 == 0
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val m = Matcher()
            __check((m(4)).toString(), "true")
            __check((m(5)).toString(), "false")
        }
