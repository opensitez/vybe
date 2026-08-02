// vybe-test: kotlin/extension_functions/test_extension_function_on_number_list_maps_all_values
// origin: languages/kotlin/tests/kotlin/test_extension_functions.rs

fun List<Int>.doubleAndSum(): Int {
            var total = 0
            for (value in this) {
                total += value * 2
            }
            return total
        }

        fun main() {
            println(listOf(1, 2, 3).doubleAndSum())
            println(listOf(10).doubleAndSum())
        }

