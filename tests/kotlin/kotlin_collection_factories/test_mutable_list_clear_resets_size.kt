// vybe-test: kotlin/kotlin_collection_factories/test_mutable_list_clear_resets_size
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_factories.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableListOf("a", "b", "c")
            values.clear()
            __check((values.isEmpty()).toString(), "true")
            values.add("x")
            __check((values.size).toString(), "1")
        }
