// vybe-test: kotlin/function_types/test_function_type_with_returning_unit_and_side_effect
// origin: languages/kotlin/tests/kotlin/test_function_types.rs

fun execute(values: List<Int>, op: (Int) -> Unit): String {
            for (v in values) op(v)
            return "done"
        }
        fun main() {
            val r = execute(listOf(1, 2)) { println("v" + it) }
            println(r)
        }

