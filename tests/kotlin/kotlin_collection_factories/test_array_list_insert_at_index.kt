// vybe-test: kotlin/kotlin_collection_factories/test_array_list_insert_at_index
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_factories.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = arrayListOf("x", "z")
            values.add(1, "y")
            __check((values.joinToString(",")).toString(), "x,y,z")
        }
