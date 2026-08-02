// vybe-test: kotlin/kotlin_iterable_to_collections/test_to_mutable_list
// origin: languages/kotlin/tests/kotlin/test_kotlin_iterable_to_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = listOf(1, 2, 3).toMutableList()
            out.add(4)
            __check((out.joinToString(",")).toString(), "1,2,3,4")
        }
