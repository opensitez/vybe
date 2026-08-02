// vybe-test: kotlin/equality_hashcode/test_structural_equality_for_data_class
// origin: languages/kotlin/tests/kotlin/test_equality_hashcode.rs

data class Item(val a: Int, val b: String)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val left = Item(1, "x")
            val right = Item(1, "x")
            val other = Item(2, "x")
            __check((left == right).toString(), "true")
            __check((left == other).toString(), "false")
        }
