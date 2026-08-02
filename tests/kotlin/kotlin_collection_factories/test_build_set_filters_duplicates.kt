// vybe-test: kotlin/kotlin_collection_factories/test_build_set_filters_duplicates
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_factories.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = buildSet {
                add(1)
                add(2)
                add(2)
                add(3)
            }
            __check((values.size).toString(), "3")
            __check((values.contains(2)).toString(), "true")
            __check((values.contains(9)).toString(), "false")
        }
