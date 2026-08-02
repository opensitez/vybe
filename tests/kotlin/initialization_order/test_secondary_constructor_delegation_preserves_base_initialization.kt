// vybe-test: kotlin/initialization_order/test_secondary_constructor_delegation_preserves_base_initialization
// origin: languages/kotlin/tests/kotlin/test_initialization_order.rs

open class Base(prefix: Int) {
            val value = prefix
            init {
                __check((value).toString(), "2")
            }
        }

        class Leaf : Base {
            val label: Int
            constructor() : this(2)
            constructor(seed: Int) : super(seed) {
                label = seed * 10
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Leaf().label).toString(), "20")
        }
