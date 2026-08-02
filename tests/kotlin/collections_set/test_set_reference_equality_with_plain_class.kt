// vybe-test: kotlin/collections_set/test_set_reference_equality_with_plain_class
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            class Label(val id: Int)
            val values = setOf(Label(1))
            __check((values.contains(Label(1))).toString(), "false")
        }
