// vybe-test: kotlin/equality_hashcode/test_list_equality_uses_element_structural_equality
// origin: languages/kotlin/tests/kotlin/test_equality_hashcode.rs

data class Item(val value: Int, val label: String)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val left = listOf(Item(1, "a"), Item(2, "b"))
            val right = listOf(Item(1, "a"), Item(2, "b"))
            __check((left == right).toString(), "true")
            __check((left[0] == right[0]).toString(), "true")
        }
