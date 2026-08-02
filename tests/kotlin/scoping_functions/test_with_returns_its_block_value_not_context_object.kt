// vybe-test: kotlin/scoping_functions/test_with_returns_its_block_value_not_context_object
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

class Holder(var count: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = with(Holder(1)) {
                count += 9
                "count:" + count
            }
            __check((out).toString(), "count:10")
        }
