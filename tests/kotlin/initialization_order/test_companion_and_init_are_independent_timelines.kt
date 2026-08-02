// vybe-test: kotlin/initialization_order/test_companion_and_init_are_independent_timelines
// origin: languages/kotlin/tests/kotlin/test_initialization_order.rs

class Holder {
            companion object {
                init { println("companion") }
            }

            init {
                println("instance")
            }
        }

        fun main() {
            Holder()
            Holder()
        }

