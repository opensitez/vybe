// vybe-test: kotlin/set_ordered_behavior/test_set_subtract_and_distinct
// origin: languages/kotlin/tests/kotlin/test_set_ordered_behavior.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = linkedSetOf(1, 2, 3, 4)
            val b = linkedSetOf(2, 4, 6)
            __check(((a - b).joinToString(",")).toString(), "1,3")
            val dup = listOf(1,1,2,2,3)
            __check((dup.toSet().joinToString(",")).toString(), "1,2,3")
        }
