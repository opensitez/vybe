// vybe-test: kotlin/initialization_order/test_companion_init_occurs_before_instance_init_once
// origin: languages/kotlin/tests/kotlin/test_initialization_order.rs

class Holder {
            companion object {
                var count = 0
                init {
                    count = 9
                }
            }

            init {
                println(companionInit())
            }
        }

        fun companionInit(): Int = Holder.count

        fun main() {
            Holder()
            Holder()
        }

