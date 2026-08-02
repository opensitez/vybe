// vybe-test: kotlin/class_delegation/test_delegation_with_custom_method_using_delegate
// origin: languages/kotlin/tests/kotlin/test_class_delegation.rs

interface Adder { fun add(a: Int): Int }

        class BaseAdder : Adder {
            override fun add(a: Int) = a + 10
        }

        class WrapperAdder(delegate: Adder) : Adder by delegate {
            fun addTwice(a: Int): Int = add(a) + add(a)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = WrapperAdder(BaseAdder())
            __check((value.addTwice(4)).toString(), "28")
        }
