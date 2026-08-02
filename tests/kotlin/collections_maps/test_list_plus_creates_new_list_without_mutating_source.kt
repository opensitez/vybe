// vybe-test: kotlin/collections_maps/test_list_plus_creates_new_list_without_mutating_source
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val head = mutableListOf(1, 2)
            val merged = head + listOf(3, 4)
            __check((merged.joinToString(",")).toString(), "1,2,3,4")
            __check((head.size).toString(), "2")
            __check((merged.size).toString(), "4")
        }
