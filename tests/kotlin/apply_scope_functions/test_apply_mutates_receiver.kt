// vybe-test: kotlin/apply_scope_functions/test_apply_mutates_receiver
// origin: languages/kotlin/tests/kotlin/test_apply_scope_functions.rs

data class Box(var value: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val box = Box(1).apply { value += 3 }
            __check((box.value).toString(), "4")
        }
