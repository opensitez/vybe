// vybe-test: kotlin/inheritance_dispatch/test_abstract_dispatch_from_chain
// origin: languages/kotlin/tests/kotlin/test_inheritance_dispatch.rs

abstract class Base {
            abstract fun emit(): Int
            open fun value(): Int = emit() * 2
        }

        class Child : Base() {
            override fun emit(): Int = 3
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val item: Base = Child()
            __check((item.value()).toString(), "6")
        }
