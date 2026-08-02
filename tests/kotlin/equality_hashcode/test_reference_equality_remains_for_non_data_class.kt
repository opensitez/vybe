// vybe-test: kotlin/equality_hashcode/test_reference_equality_remains_for_non_data_class
// origin: languages/kotlin/tests/kotlin/test_equality_hashcode.rs

class Item(val id: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val first = Item(1)
            val second = first
            val third = Item(1)
            __check((first == third).toString(), "false")
            __check((first === second).toString(), "true")
            __check((first === third).toString(), "false")
        }
