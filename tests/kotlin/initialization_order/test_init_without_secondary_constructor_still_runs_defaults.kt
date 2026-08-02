// vybe-test: kotlin/initialization_order/test_init_without_secondary_constructor_still_runs_defaults
// origin: languages/kotlin/tests/kotlin/test_initialization_order.rs

class Holder(val base: Int = 1) {
            val scaled = base * 2
            init { println(scaled) }
        }

        fun main() {
            Holder()
            Holder(4)
        }

