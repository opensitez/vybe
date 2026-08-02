// vybe-test: kotlin/scoping_functions/test_with_supports_reassigning_outer_mutable_state
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

class Holder(var value: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var mutable = Holder(4)
            val label = with(mutable) {
                value *= 3
                "v" + value
            }
            __check((label).toString(), "v12")
            __check((mutable.value).toString(), "12")
        }
