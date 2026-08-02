// vybe-test: kotlin/array_indexing/test_nested_list_mutation_by_index
// origin: languages/kotlin/tests/kotlin/test_array_indexing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val m = mutableListOf(mutableListOf(1,2), mutableListOf(3,4))
m[1][0] = 9
__check((m[1][0]).toString(), "9") }
