// vybe-test: kotlin/apply_scope_functions/test_also_preserves_receiver_and_side_effects
// origin: languages/kotlin/tests/kotlin/test_apply_scope_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = mutableListOf<Int>()
            val values = intArrayOf(1, 2, 3).toMutableList().also {
                out.addAll(it)
                it.add(4)
            }
            __check((values.joinToString(",")).toString(), "1,2,3,4")
            __check((out.joinToString(",")).toString(), "1,2,3")
        }
