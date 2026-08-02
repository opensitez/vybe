// vybe-test: kotlin/equality_hashcode/test_equals_with_null_and_different_type
// origin: languages/kotlin/tests/kotlin/test_equality_hashcode.rs

data class Item(val id: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val item = Item(1)
            __check((item == null).toString(), "false")
            __check((item == 1).toString(), "false")
        }
