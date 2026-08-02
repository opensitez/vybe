// vybe-test: kotlin/kotlin_collection_factories/test_reified_builders_do_not_escape_mutation
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_factories.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base = mutableListOf(1)
            val built = buildList {
                addAll(base)
                add(2)
            }
            base.add(9)
            __check((base.joinToString(",")).toString(), "1,9")
            __check((built.joinToString(",")).toString(), "1,2")
        }
