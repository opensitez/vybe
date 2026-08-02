// vybe-test: kotlin/function_types/test_function_type_with_boolean_predicate_list
// origin: languages/kotlin/tests/kotlin/test_function_types.rs

fun filter(values: List<Int>, keep: (Int) -> Boolean): List<Int> {
            return values.filter(keep)
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = filter(listOf(1, 2, 3, 4), { it % 2 == 0 })
            __check((out.joinToString(",")).toString(), "2,4")
        }
