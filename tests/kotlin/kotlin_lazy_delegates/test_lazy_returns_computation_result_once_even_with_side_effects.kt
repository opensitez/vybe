// vybe-test: kotlin/kotlin_lazy_delegates/test_lazy_returns_computation_result_once_even_with_side_effects
// origin: languages/kotlin/tests/kotlin/test_kotlin_lazy_delegates.rs

var sideEffect = 0
        fun load(): Int {
            sideEffect += 2
            return sideEffect
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value by lazy { load() }
            __check((value).toString(), "2")
            __check((value).toString(), "2")
            __check((sideEffect).toString(), "2")
        }
