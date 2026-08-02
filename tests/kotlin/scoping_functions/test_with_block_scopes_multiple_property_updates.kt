// vybe-test: kotlin/scoping_functions/test_with_block_scopes_multiple_property_updates
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

class Holder {
            var value = 1
            fun add(step: Int) { value += step }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = with(Holder()) {
                add(3)
                add(2)
                value
            }
            __check((out).toString(), "6")
        }
