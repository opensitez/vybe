// vybe-test: kotlin/data_classes/test_data_class_in_set_uses_equality
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Item(val id: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val first = Item(1)
            val second = Item(1)
            val values = setOf(first)
            __check((values.contains(second)).toString(), "true")
        }
