// vybe-test: kotlin/collections_iterables/test_associate_by_last_duplicate_key_wins
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val entries = listOf("a" to 1, "b" to 2, "a" to 9)
            val map = entries.associateBy({ it.first }) { it.second }
            __check((map["a"]).toString(), "9")
            __check((map["b"]).toString(), "2")
            __check((map.size).toString(), "2")
        }
