// vybe-test: kotlin/kotlin_list_mutation_order/test_mutable_list_basic_updates
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_mutation_order.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val l = mutableListOf(1, 3, 4)
            l.add(5)
            l[1] = 2
            __check((l.toString()).toString(), "[1, 2, 4, 5]")
        }
