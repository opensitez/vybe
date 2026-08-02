// vybe-test: kotlin/scoping_functions/test_nested_scoping_functions_show_receiver_visibility
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

class Counter(var value: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = Counter(1).apply {
                this.value += 1
                val local = Counter(value).also {
                    it.value += 4
                }
                this.value += local.value
            }
            __check((out.value).toString(), "7")
        }
