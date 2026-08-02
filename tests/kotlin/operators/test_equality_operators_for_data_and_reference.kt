// vybe-test: kotlin/operators/test_equality_operators_for_data_and_reference
// origin: languages/kotlin/tests/kotlin/test_operators.rs

class Cell(val value: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Cell(1)
            val b = Cell(1)
            val c = a
            __check((a == b).toString(), "true")
            __check((a === c).toString(), "true")
            __check((a === b).toString(), "false")
        }
