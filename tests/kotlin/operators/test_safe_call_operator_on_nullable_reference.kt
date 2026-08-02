// vybe-test: kotlin/operators/test_safe_call_operator_on_nullable_reference
// origin: languages/kotlin/tests/kotlin/test_operators.rs

class Holder(val label: String)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val absent: Holder? = null
            val present: Holder? = Holder("value")
            __check((absent?.label).toString(), "null")
            __check((present?.label).toString(), "value")
        }
