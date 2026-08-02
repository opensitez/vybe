// vybe-test: kotlin/initialization_order/test_init_uses_global_counter_and_reuses_updated_state
// origin: languages/kotlin/tests/kotlin/test_initialization_order.rs

var stamp = 0

        fun next_stamp(): Int {
            stamp += 1
            return stamp
        }

        class Holder {
            val first = next_stamp()
            val second = first + 10
            init {
                println(first)
                println(second)
            }
            val third = next_stamp()
            init {
                println(third)
            }
        }

        fun main() {
            Holder()
            Holder()
            println(stamp)
        }

