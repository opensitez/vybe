// vybe-test: kotlin/kotlin_list_mutation_order/test_mutable_list_remove_first
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_mutation_order.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val l = mutableListOf("x", "y", "z")
            l.removeAt(0)
            __check((l.toString()).toString(), "[y, z]")
        }
