// vybe-test: kotlin/properties/test_property_accessor_reacts_to_mutating_inputs
// origin: languages/kotlin/tests/kotlin/test_properties.rs

class ScoreBoard {
            var raw = 3
            val rating: Int
                get() = raw * 2
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val board = ScoreBoard()
            board.raw = 4
            __check((board.rating).toString(), "8")
        }
