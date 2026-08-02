// vybe-test: kotlin/invoke_operator/test_invoke_chainable
// origin: languages/kotlin/tests/kotlin/test_invoke_operator.rs

class A {
            operator fun invoke(x: Int): B = B(x + 1)
        }
        class B(val v: Int) {
            operator fun invoke(y: Int): Int = v + y
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((A()(3)(4)).toString(), "8")
        }
