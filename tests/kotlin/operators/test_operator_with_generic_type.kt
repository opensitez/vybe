// vybe-test: kotlin/operators/test_operator_with_generic_type
// origin: languages/kotlin/tests/kotlin/test_operators.rs

class Box<T>(private val value: T) {
            operator fun plus(other: Box<T>): String {
                return this.value.toString() + other.value.toString()
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Box("a") + Box("b")).toString(), "ab")
            __check((Box(1) + Box(2)).toString(), "12")
        }
