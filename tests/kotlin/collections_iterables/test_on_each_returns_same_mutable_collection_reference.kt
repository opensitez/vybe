// vybe-test: kotlin/collections_iterables/test_on_each_returns_same_mutable_collection_reference
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = mutableListOf(1, 2, 3)
            val same = nums.onEach { }
            same[0] = 9
            __check((nums[0]).toString(), "9")
            __check((nums === same).toString(), "true")
        }
