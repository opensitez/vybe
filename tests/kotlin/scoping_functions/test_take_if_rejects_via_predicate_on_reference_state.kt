// vybe-test: kotlin/scoping_functions/test_take_if_rejects_via_predicate_on_reference_state
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

class Box(var n: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = Box(2)
            val filtered = value.takeIf { it.n > 2 }
            __check((filtered == null).toString(), "true")
            __check((value.n).toString(), "2")
        }
