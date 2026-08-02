// vybe-test: kotlin/equality_hashcode/test_reference_equality_still_work_for_same_and_different_instances
// origin: languages/kotlin/tests/kotlin/test_equality_hashcode.rs

class Item(val label: String)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val first = Item("x")
            val second = Item("x")
            val same = first
            __check((first === first).toString(), "true")
            __check((first === same).toString(), "true")
            __check((first === second).toString(), "false")
        }
