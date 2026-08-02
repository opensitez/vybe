// vybe-test: kotlin/kotlin_collection_factories/test_hash_map_contains_checks_both_key_and_value
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_factories.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = hashMapOf("x" to 10, "y" to 20)
            __check((value.containsKey("x")).toString(), "true")
            __check((value.containsValue(20)).toString(), "true")
            __check((value.containsValue(30)).toString(), "false")
        }
