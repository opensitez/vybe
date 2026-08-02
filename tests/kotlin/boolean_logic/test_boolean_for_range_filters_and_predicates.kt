// vybe-test: kotlin/boolean_logic/test_boolean_for_range_filters_and_predicates
// origin: languages/kotlin/tests/kotlin/test_boolean_logic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = intArrayOf(1, 2, 3, 4, 5)
            val evens = nums.filter { it % 2 == 0 }
            val odds = nums.filter { it % 2 != 0 }
            __check((evens.joinToString(",")).toString(), "2,4")
            __check((odds.joinToString(",")).toString(), "1,3,5")
        }
