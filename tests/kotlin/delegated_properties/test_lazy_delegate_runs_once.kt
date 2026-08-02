// vybe-test: kotlin/delegated_properties/test_lazy_delegate_runs_once
// origin: languages/kotlin/tests/kotlin/test_delegated_properties.rs

class Counter {
            var hits = 0
            val value by lazy { hits += 1
42 }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val counter = Counter()
            __check((counter.value).toString(), "42")
            __check((counter.value).toString(), "42")
            __check((counter.hits).toString(), "1")
        }
