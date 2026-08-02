// vybe-test: kotlin/invoke_operator/test_invoke_operator_variance_through_type_alias
// origin: languages/kotlin/tests/kotlin/test_invoke_operator.rs

typealias IntCall = (Int) -> Int
        class Wrapper {
            operator fun invoke(v: Int): IntCall = { it + v }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val f = Wrapper()(3)
            __check((f(7)).toString(), "10")
        }
