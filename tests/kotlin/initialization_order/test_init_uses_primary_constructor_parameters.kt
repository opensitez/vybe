// vybe-test: kotlin/initialization_order/test_init_uses_primary_constructor_parameters
// origin: languages/kotlin/tests/kotlin/test_initialization_order.rs

class Holder(prefix: String) {
            val value: String

            init {
                value = prefix.uppercase()
                println(value)
            }
        }

        fun main() {
            Holder("x")
            Holder("y")
        }

