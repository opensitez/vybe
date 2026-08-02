// vybe-test: kotlin/member_references/test_reference_to_nested_member_function
// origin: languages/kotlin/tests/kotlin/test_member_references.rs

class Board {
            class Cell {
                fun mark(v: String): String = "[$v]"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val label = Board.Cell::mark
            __check((label(Board.Cell(), "x")).toString(), "[x]")
        }
