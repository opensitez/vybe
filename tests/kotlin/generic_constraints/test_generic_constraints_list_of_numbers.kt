// vybe-test: kotlin/generic_constraints/test_generic_constraints_list_of_numbers
// origin: languages/kotlin/tests/kotlin/test_generic_constraints.rs

fun <T : Number> total(values: List<T>): Int {
            var n = 0
            for (v in values) n += v.toInt()
            return n
        }
        fun main() {
            println(total(listOf(1, 2, 3)))
            println(total(listOf(1.2, 2.8)))
        }

