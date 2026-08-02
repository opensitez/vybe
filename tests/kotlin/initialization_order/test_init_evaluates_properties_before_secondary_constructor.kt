// vybe-test: kotlin/initialization_order/test_init_evaluates_properties_before_secondary_constructor
// origin: languages/kotlin/tests/kotlin/test_initialization_order.rs

class Holder(val value: Int) {
            val label: String

            init {
                label = "v=" + value.toString()
            }

            constructor() : this(5)

            fun out(): String = label
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val item = Holder()
            __check((item.out()).toString(), "v=5")
        }
