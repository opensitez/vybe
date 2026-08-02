// vybe-test: kotlin/collections_set/test_set_union_intersection_do_not_mutate_operands
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = linkedSetOf(1, 2, 3)
            val b = setOf(3, 4)
            val union = a union b
            val inter = a intersect b
            __check((union.size).toString(), "4")
            __check((inter.size).toString(), "1")
            __check((a.size).toString(), "3")
            __check((b.size).toString(), "2")
            __check((a.contains(4)).toString(), "false")
            __check((b.contains(1)).toString(), "false")
        }
