// vybe-test: kotlin/collections_set/test_set_mapping_changes_type
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = setOf(1, 2, 3, 4)
            val mapped = values.map { it * 2 }
            val restored = mapped.toSet()
            __check((mapped.size).toString(), "4")
            __check((restored.size).toString(), "4")
            __check((restored.contains(8)).toString(), "true")
        }
