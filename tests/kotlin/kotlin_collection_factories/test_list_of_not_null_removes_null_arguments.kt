// vybe-test: kotlin/kotlin_collection_factories/test_list_of_not_null_removes_null_arguments
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_factories.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOfNotNull(1, null, 2, null, 3)
            __check((values.size).toString(), "3")
            __check((values.joinToString(",")).toString(), "1,2,3")
        }
