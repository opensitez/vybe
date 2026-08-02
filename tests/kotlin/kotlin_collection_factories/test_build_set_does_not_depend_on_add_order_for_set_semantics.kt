// vybe-test: kotlin/kotlin_collection_factories/test_build_set_does_not_depend_on_add_order_for_set_semantics
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_factories.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = buildSet {
                add("b")
                add("a")
                add("a")
                add("c")
            }
            __check((values.contains("a")).toString(), "true")
            __check((values.contains("c")).toString(), "true")
            __check((values.size).toString(), "3")
        }
