// vybe-test: kotlin/inheritance_dispatch/test_generic_inheritance_dispatch_on_bounds
// origin: languages/kotlin/tests/kotlin/test_inheritance_dispatch.rs

interface ValueCarrier {
            fun value(): Int
        }

        open class Base<T : ValueCarrier> : ValueCarrier {
            override fun value(): Int = 0
        }

        class Child : Base<Node>() {
            override fun value(): Int = 7
        }

        class Node : ValueCarrier {
            override fun value(): Int = 2
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val item: Base<*> = Child()
            val direct: ValueCarrier = Child()
            __check((item.value()).toString(), "7")
            __check((direct.value()).toString(), "7")
        }
