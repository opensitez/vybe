// vybe-test: kotlin/kotlin_lazy_initialization/test_lazy_value_can_depend_on_previous_state
// origin: languages/kotlin/tests/kotlin/test_kotlin_lazy_initialization.rs

class Holder {
            var seed = 5
            val value: Int by lazy {
                seed * 2
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val h = Holder()
            __check((h.value).toString(), "10")
            h.seed = 9
            __check((h.value).toString(), "10")
        }
