// vybe-test: kotlin/kotlin_collection_factories/test_hash_map_allows_null_value_storage
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_factories.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = HashMap<String, String?>()
            value["a"] = null
            value["b"] = "ok"
            __check((value["a"] == null).toString(), "true")
            __check((value["b"]).toString(), "ok")
            __check((value.size).toString(), "2")
        }
