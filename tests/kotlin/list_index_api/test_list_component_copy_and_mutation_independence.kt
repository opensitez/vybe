// vybe-test: kotlin/list_index_api/test_list_component_copy_and_mutation_independence
// origin: languages/kotlin/tests/kotlin/test_list_index_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = mutableListOf(1, 2, 3)
            val copy = source.toMutableList()
            copy.add(4)
            __check((source.size).toString(), "3")
            __check((copy.size).toString(), "4")
            __check((source.joinToString(",")).toString(), "1,2,3")
            __check((copy.joinToString(",")).toString(), "1,2,3,4")
        }
