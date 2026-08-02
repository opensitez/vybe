// vybe-test: kotlin/collections/test_array_to_mutable_list_and_mutation_contract
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = arrayOf(1, 2, 3)
            val mutable = nums.toMutableList()
            mutable.add(4)
            mutable[0] = 9
            __check((nums.joinToString(",")).toString(), "1,2,3")
            __check((mutable.joinToString(",")).toString(), "9,2,3,4")
        }
