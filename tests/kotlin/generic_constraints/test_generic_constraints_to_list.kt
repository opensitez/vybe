// vybe-test: kotlin/generic_constraints/test_generic_constraints_to_list
// origin: languages/kotlin/tests/kotlin/test_generic_constraints.rs

fun <T> toFlat(values: List<List<T>>): List<T> = values.flatten()
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = toFlat(listOf(listOf(1, 2), listOf(3, 4)))
            __check((out.joinToString(",")).toString(), "1,2,3,4")
        }
