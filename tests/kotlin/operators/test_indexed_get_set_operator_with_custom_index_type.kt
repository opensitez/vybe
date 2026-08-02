// vybe-test: kotlin/operators/test_indexed_get_set_operator_with_custom_index_type
// origin: languages/kotlin/tests/kotlin/test_operators.rs

class Slots {
            private val values = arrayOf(9, 8, 7)
            operator fun get(index: Int): Int = values[index]
            operator fun set(index: Int, value: Int) {
                values[index] = value
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val box = Slots()
            __check((box[0] + box[2]).toString(), "16")
            box[1] = 4
            __check((box[1]).toString(), "4")
        }
