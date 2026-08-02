// vybe-test: kotlin/kotlin_collection_factories/test_empty_list_of_not_null
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_factories.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOfNotNull<Int>()
            __check((values.isEmpty()).toString(), "true")
            __check((values.size).toString(), "0")
        }
