// vybe-test: kotlin/kotlin_pairs_triples/test_map_entry_as_pair_api
// origin: languages/kotlin/tests/kotlin/test_kotlin_pairs_triples.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val items = mapOf("a" to 1, "b" to 2)
            val first = items.entries.first()
            __check((first.key).toString(), "a")
            __check((first.value).toString(), "1")
        }
