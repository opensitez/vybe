// vybe-test: kotlin/collections_maps/test_list_reversed_returns_independent_copy
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val original = mutableListOf(1, 2, 3)
            val reversed = original.reversed()
            __check((reversed.joinToString(",")).toString(), "3,2,1")
            original[0] = 9
            __check((reversed.joinToString(",")).toString(), "3,2,1")
            __check((original.joinToString(",")).toString(), "9,2,3")
        }
