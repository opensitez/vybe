// vybe-test: kotlin/equality_hashcode/test_equals_with_different_types_is_false
// origin: languages/kotlin/tests/kotlin/test_equality_hashcode.rs

data class Item(val value: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val item = Item(1)
            __check((item == 1).toString(), "false")
            __check((item.equals("x")).toString(), "false")
        }
