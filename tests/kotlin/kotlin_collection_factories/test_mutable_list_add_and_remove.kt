// vybe-test: kotlin/kotlin_collection_factories/test_mutable_list_add_and_remove
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_factories.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableListOf(1, 2)
            values.add(3)
            values.removeAt(1)
            values.add(1, 8)
            __check((values.joinToString(",")).toString(), "1,8,3")
        }
