// vybe-test: kotlin/data_classes/test_data_class_in_map_for_hash_lookup_after_copy
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Key(val id: Int, val label: String)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val original = Key(1, "a")
            val map = mutableMapOf(original to "first")
            val copy = original.copy()
            original.label = "b"
            __check((map[original] == null).toString(), "true")
            __check((map[copy]).toString(), "first")
        }
