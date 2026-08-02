// vybe-test: kotlin/kotlin_iterable_to_collections/test_to_list_from_array
// origin: languages/kotlin/tests/kotlin/test_kotlin_iterable_to_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val arr = arrayOf(1, 2, 3)
            val out = arr.asList()
            __check((out.joinToString(",")).toString(), "1,2,3")
        }
