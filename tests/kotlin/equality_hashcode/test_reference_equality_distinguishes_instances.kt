// vybe-test: kotlin/equality_hashcode/test_reference_equality_distinguishes_instances
// origin: languages/kotlin/tests/kotlin/test_equality_hashcode.rs

class Holder(val value: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val left = Holder(1)
            val right = Holder(1)
            __check((left === right).toString(), "false")
            __check((left === left).toString(), "true")
        }
