// vybe-test: kotlin/ordered_collections/test_set_to_string_preserves_deterministic_for_linked_set
// origin: languages/kotlin/tests/kotlin/test_ordered_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val set = linkedSetOf("x", "y", "z")
            __check((set.toString()).toString(), "[x, y, z]")
        }
