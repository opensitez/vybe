// vybe-test: kotlin/kotlin_collection_factories/test_build_list_collects_all_items
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_factories.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = buildList {
                add(1)
                add(2)
                add(3)
            }
            __check((values.size).toString(), "3")
            __check((values.joinToString(",")).toString(), "1,2,3")
        }
