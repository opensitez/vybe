// vybe-test: kotlin/scoping_functions/test_apply_on_custom_type
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

class Counter {
            var value: Int = 0
            fun bump(step: Int) { value += step }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val counter = Counter().apply {
                bump(2)
                bump(3)
            }
            __check((counter.value).toString(), "5")
        }
