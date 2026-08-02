// vybe-test: kotlin/kotlin_collection_factories/test_list_of_not_null_from_expressions
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_factories.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOfNotNull(if (true) "on" else null, null, if (false) "no" else "yes")
            __check((values.joinToString("-")).toString(), "on-yes")
        }
