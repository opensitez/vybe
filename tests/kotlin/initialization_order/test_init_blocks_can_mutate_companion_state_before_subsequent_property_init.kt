// vybe-test: kotlin/initialization_order/test_init_blocks_can_mutate_companion_state_before_subsequent_property_init
// origin: languages/kotlin/tests/kotlin/test_initialization_order.rs

var globalValue = 1

        class Holder {
            val first = globalValue

            init {
                globalValue = 10
            }

            val second = globalValue * 2

            init {
                println(first)
                println(second)
            }
        }

        fun main() {
            Holder()
            Holder()
            println(globalValue)
        }

