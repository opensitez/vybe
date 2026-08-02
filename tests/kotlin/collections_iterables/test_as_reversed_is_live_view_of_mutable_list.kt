// vybe-test: kotlin/collections_iterables/test_as_reversed_is_live_view_of_mutable_list
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = mutableListOf(1, 2, 3)
            val view = nums.asReversed()
            __check((view.joinToString(",")).toString(), "3,2,1")
            nums.add(4)
            __check((view.joinToString(",")).toString(), "4,3,2,1")
            view[0] = 10
            __check((nums.joinToString(",")).toString(), "10,2,3,1,4")
        }
