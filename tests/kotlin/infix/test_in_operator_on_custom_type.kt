// vybe-test: kotlin/infix/test_in_operator_on_custom_type
// origin: languages/kotlin/tests/kotlin/test_infix.rs

class Bag {
            val values = arrayOf(1, 2, 3)
            fun has(value: Int): Boolean {
                return values[0] == value || values[1] == value || values[2] == value
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val b = Bag()
            if (2 in 1..4) {
                __check(("range").toString(), "range")
            }
            if (b.has(2)) {
                __check(("found").toString(), "found")
            }
        }
