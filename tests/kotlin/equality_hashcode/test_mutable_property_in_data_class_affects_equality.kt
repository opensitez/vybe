// vybe-test: kotlin/equality_hashcode/test_mutable_property_in_data_class_affects_equality
// origin: languages/kotlin/tests/kotlin/test_equality_hashcode.rs

data class Item(var value: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val left = Item(1)
            val set = hashSetOf(left)
            left.value = 2
            __check((set.contains(Item(1))).toString(), "false")
            __check((set.contains(Item(2))).toString(), "false")
        }
