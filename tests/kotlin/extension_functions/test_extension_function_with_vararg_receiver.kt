// vybe-test: kotlin/extension_functions/test_extension_function_with_vararg_receiver
// origin: languages/kotlin/tests/kotlin/test_extension_functions.rs

fun Int.joinWith(vararg values: Int): String {
            var total = this
            for (value in values) {
                total += value
            }
            return total.toString()
        }

        fun main() {
            println(1.joinWith(2, 3, 4))
            println(0.joinWith())
        }

