// vybe-test: kotlin/kotlin_lazy_initialization/test_lazy_initialized_property_runs_once
// origin: languages/kotlin/tests/kotlin/test_kotlin_lazy_initialization.rs

class Holder {
            var calls = 0
            val value: Int by lazy {
                calls += 1
                10
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
            __check((h.calls).toString(), "0")
            __check((h.value).toString(), "10")
            __check((h.value).toString(), "10")
            __check((h.calls).toString(), "1")
        }
