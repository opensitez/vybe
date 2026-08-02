// vybe-test: kotlin/collection_projection_views/test_iterator_mutability_contracts
// origin: languages/kotlin/tests/kotlin/test_collection_projection_views.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableListOf(1, 2, 3)
            val it = values.listIterator()
            __check((it.next()).toString(), "1")
            it.set(7)
            __check((values.joinToString(",")).toString(), "7,2,3")
            it.add(9)
            __check((values.joinToString(",")).toString(), "7,9,2,3")
            it.previous()
            it.remove()
            __check((values.joinToString(",")).toString(), "7,2,3")
        }
