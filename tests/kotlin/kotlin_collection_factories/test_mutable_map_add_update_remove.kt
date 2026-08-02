// vybe-test: kotlin/kotlin_collection_factories/test_mutable_map_add_update_remove
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_factories.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = mutableMapOf("x" to 1)
            value["y"] = 2
            value["x"] = 9
            value.remove("y")
            __check((value["x"]).toString(), "9")
            __check((value.containsKey("y")).toString(), "false")
            __check((value.size).toString(), "1")
        }
