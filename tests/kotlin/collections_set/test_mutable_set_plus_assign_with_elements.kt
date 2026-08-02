// vybe-test: kotlin/collections_set/test_mutable_set_plus_assign_with_elements
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base = mutableSetOf(1, 2)
            val snapshot = base.toSet()
            base += setOf(2, 3, 4)
            __check((base.size).toString(), "4")
            __check((snapshot.size).toString(), "2")
            __check((base.contains(4)).toString(), "true")
        }
