// vybe-test: kotlin/collections_maps/test_list_creation_and_indexing
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = listOf(10, 20, 30)
            __check((nums.size).toString(), "3")
            __check((nums[0]).toString(), "10")
            __check((nums[2]).toString(), "30")
        }
