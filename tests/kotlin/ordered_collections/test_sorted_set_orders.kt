// vybe-test: kotlin/ordered_collections/test_sorted_set_orders
// origin: languages/kotlin/tests/kotlin/test_ordered_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val set = java.util.TreeSet<Int>()
            set.add(3)
            set.add(1)
            set.add(2)
            __check((set.joinToString(",")).toString(), "1,2,3")
        }
