// vybe-test: kotlin/collections_iterables/test_list_associate_by_and_group_by
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val people = listOf(
                Pair("alice", "A"),
                Pair("bob", "A"),
                Pair("charlie", "B")
            )
            val byGroup = people.groupBy { it.second }
            __check((byGroup["A"]?.size ?: 0).toString(), "2")
            val mapByName = people.associateBy { it.first }
            __check((mapByName["charlie"]?.second).toString(), "B")
        }
