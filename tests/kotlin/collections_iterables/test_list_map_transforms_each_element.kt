// vybe-test: kotlin/collections_iterables/test_list_map_transforms_each_element
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = listOf(1, 2, 3)
            val doubled = nums.map { it * 2 }
            __check((doubled.joinToString(",")).toString(), "2,4,6")
        }
