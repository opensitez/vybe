// vybe-test: kotlin/operators/test_custom_index_get_set
// origin: languages/kotlin/tests/kotlin/test_operators.rs

class Buckets {
            private val data = arrayOf(5, 10, 15)
            operator fun get(index: Int): Int {
                return data[index]
            }
            operator fun set(index: Int, value: Int) {
                data[index] = value
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val storage = Buckets()
            __check((storage[0]).toString(), "5")
            storage[1] = 25
            __check((storage[1]).toString(), "25")
            __check((storage[2]).toString(), "15")
        }
