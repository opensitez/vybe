// vybe-test: kotlin/kotlin_collection_factories/test_linked_map_preserves_insertion_order
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_factories.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = linkedMapOf("first" to 1, "second" to 2, "third" to 3)
            __check((value.keys.joinToString(",")).toString(), "first,second,third")
            value["second"] = 7
            __check((value.keys.joinToString(",")).toString(), "first,second,third")
        }
