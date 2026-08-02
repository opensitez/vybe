// vybe-test: kotlin/kotlin_collection_factories/test_array_list_is_mutable_collection
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_factories.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = ArrayList<Int>()
            values.add(9)
            values.add(7)
            values.add(5)
            values[1] = 4
            __check((values.joinToString(",")).toString(), "9,4,5")
        }
