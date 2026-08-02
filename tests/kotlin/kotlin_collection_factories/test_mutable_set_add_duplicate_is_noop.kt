// vybe-test: kotlin/kotlin_collection_factories/test_mutable_set_add_duplicate_is_noop
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_factories.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableSetOf("a")
            val first = values.add("a")
            val second = values.add("b")
            __check((first).toString(), "false")
            __check((second).toString(), "true")
            __check((values.size).toString(), "2")
        }
