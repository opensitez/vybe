// vybe-test: kotlin/collections_iterables/test_list_of_not_null_omits_empty_inputs
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOfNotNull(null, 7, null, 0, 4)
            __check((values.size).toString(), "3")
            __check((values.joinToString(",")).toString(), "7,0,4")
        }
