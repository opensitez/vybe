// vybe-test: kotlin/extension_functions/test_extension_function_on_int_array_returns_product
// origin: languages/kotlin/tests/kotlin/test_extension_functions.rs

fun IntArray.product(): Int {
            var total = 1
            for (value in this) {
                total *= value
            }
            return total
        }

        fun main() {
            println(intArrayOf(2, 3, 4).product())
        }

