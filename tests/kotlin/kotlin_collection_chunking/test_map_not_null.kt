// vybe-test: kotlin/kotlin_collection_chunking/test_map_not_null
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_chunking.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf("1", null, "x", "3")
            val out = values.mapNotNull { it?.toIntOrNull() }
            __check((out.joinToString(",")).toString(), "1,3")
        }
