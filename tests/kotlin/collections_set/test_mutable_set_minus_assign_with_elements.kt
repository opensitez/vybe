// vybe-test: kotlin/collections_set/test_mutable_set_minus_assign_with_elements
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableSetOf(1, 2, 3, 4)
            values -= setOf(2, 9, 4)
            __check((values.size).toString(), "2")
            __check((values.contains(2)).toString(), "false")
            __check((values.contains(4)).toString(), "false")
            __check((values.contains(3)).toString(), "true")
        }
