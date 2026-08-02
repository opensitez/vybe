// vybe-test: kotlin/kotlin_collection_factories/test_build_list_size_hint_affects_internal_growth
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_factories.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = buildList(2) {
                add("a")
                add("b")
                add("c")
            }
            __check((values.size).toString(), "3")
            __check((values.joinToString(":")).toString(), "a:b:c")
        }
