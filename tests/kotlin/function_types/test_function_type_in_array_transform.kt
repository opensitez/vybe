// vybe-test: kotlin/function_types/test_function_type_in_array_transform
// origin: languages/kotlin/tests/kotlin/test_function_types.rs

fun main() {
            val ops: Array<(Int) -> Int> = arrayOf({ it + 1 }, { it - 1 })
            var value = 10
            for (op in ops) {
                value = op(value)
            }
            println(value)
        }

