// vybe-test: kotlin/operators/test_custom_inc_dec_operators
// origin: languages/kotlin/tests/kotlin/test_operators.rs

class Counter(var value: Int) {
            operator fun inc(): Counter {
                value += 1
                return this
            }

            operator fun dec(): Counter {
                value -= 1
                return this
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var counter = Counter(2)
            counter++
            __check((counter.value).toString(), "3")
            counter--
            __check((counter.value).toString(), "2")
        }
