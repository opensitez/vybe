// vybe-test: kotlin/operators/test_custom_plus_minus_operators
// origin: languages/kotlin/tests/kotlin/test_operators.rs

class Counter(val value: Int) {
            operator fun plus(other: Counter): Counter = Counter(value + other.value)
            operator fun minus(other: Counter): Counter = Counter(value - other.value)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Counter(10)
            val b = Counter(4)
            __check(((a + b).value).toString(), "14")
            __check(((a - b).value).toString(), "6")
        }
