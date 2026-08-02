// vybe-test: kotlin/collections_set/test_set_replace_element_after_remove_by_equal_shape
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

data class Box(val id: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableSetOf(Box(1), Box(2))
            values.remove(Box(1))
            values.add(Box(3))
            __check((values.size).toString(), "2")
            __check((values.contains(Box(2))).toString(), "true")
            __check((values.contains(Box(1))).toString(), "false")
            __check((values.contains(Box(3))).toString(), "true")
        }
