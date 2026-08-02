// vybe-test: kotlin/kotlin_collection_factories/test_map_of_single_entry_lookup
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_factories.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = mapOf("a" to 1)
            __check((value["a"]).toString(), "1")
            __check((value["b"] ?: -1).toString(), "-1")
        }
