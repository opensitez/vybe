// vybe-test: kotlin/class_delegation/test_delegate_class_for_multiple_calls
// origin: languages/kotlin/tests/kotlin/test_class_delegation.rs

interface Calc { fun step(v: Int): Int }

        class Base : Calc {
            override fun step(v: Int) = v + 1
        }

        class Wrapper(private val base: Calc) : Calc by base

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = Wrapper(Base())
            __check((c.step(0)).toString(), "1")
            __check((c.step(9)).toString(), "10")
        }
