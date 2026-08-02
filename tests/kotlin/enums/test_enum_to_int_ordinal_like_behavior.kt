// vybe-test: kotlin/enums/test_enum_to_int_ordinal_like_behavior
// origin: languages/kotlin/tests/kotlin/test_enums.rs

enum class Dice {
            ONE, TWO, THREE
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Dice.ONE
            val b = Dice.TWO
            val c = Dice.THREE
            __check((a).toString(), "0")
            __check((c).toString(), "2")
        }
