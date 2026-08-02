// vybe-test: kotlin/kotlin_lazy_delegates/test_lazy_after_mutation_from_other_property
// origin: languages/kotlin/tests/kotlin/test_kotlin_lazy_delegates.rs

class Holder {
            var seed = 1
            val value by lazy { seed * 10 }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val h = Holder()
            h.seed = 3
            __check((h.value).toString(), "30")
            h.seed = 9
            __check((h.value).toString(), "30")
        }
