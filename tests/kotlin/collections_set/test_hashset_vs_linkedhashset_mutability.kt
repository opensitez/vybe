// vybe-test: kotlin/collections_set/test_hashset_vs_linkedhashset_mutability
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val linked = linkedSetOf(1, 2, 3)
            val hash = hashSetOf(3, 2, 1)
            __check((linked.size).toString(), "3")
            __check((hash.size).toString(), "3")
            __check(((linked - setOf(2)).contains(2)).toString(), "false")
            __check(((hash + setOf(4)).contains(4)).toString(), "true")
        }
