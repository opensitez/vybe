// vybe-test: kotlin/apply_scope_functions/test_with_configures_receiver_without_reference
// origin: languages/kotlin/tests/kotlin/test_apply_scope_functions.rs

data class Accumulator(var total: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val acc = Accumulator(0).apply {
                total += 1
                total += 2
            }
            val done = with(acc) {
                total += 5
                total
            }
            __check((acc.total).toString(), "8")
            __check((done).toString(), "8")
        }
