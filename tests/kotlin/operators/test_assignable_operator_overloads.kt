// vybe-test: kotlin/operators/test_assignable_operator_overloads
// origin: languages/kotlin/tests/kotlin/test_operators.rs

class Counter(var value: Int) {
            operator fun plusAssign(other: Counter) {
                value += other.value
            }

            operator fun minusAssign(other: Counter) {
                value -= other.value
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val acc = Counter(10)
            acc += Counter(3)
            __check((acc.value).toString(), "13")
            acc -= Counter(1)
            __check((acc.value).toString(), "12")
        }
