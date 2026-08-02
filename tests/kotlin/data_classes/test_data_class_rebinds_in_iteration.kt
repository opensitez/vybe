// vybe-test: kotlin/data_classes/test_data_class_rebinds_in_iteration
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Meter(val id: Int, val value: Int)

        fun main() {
            val items = mutableListOf(Meter(1, 1), Meter(2, 2))
            var sum = 0
            for (item in items) {
                val updated = item.copy(value = item.value + 5)
                sum += updated.value
            }
            println(sum)
            println(items[0].value)
        }

