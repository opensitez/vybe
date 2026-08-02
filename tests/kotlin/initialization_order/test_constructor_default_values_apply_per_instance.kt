// vybe-test: kotlin/initialization_order/test_constructor_default_values_apply_per_instance
// origin: languages/kotlin/tests/kotlin/test_initialization_order.rs

class Holder(prefix: Int = 1) {
            val value = prefix + 1
            init {
                println(value)
            }
        }

        fun main() {
            Holder()
            Holder(2)
        }

