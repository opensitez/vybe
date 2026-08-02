// vybe-test: kotlin/delegated_properties/test_lazy_reuses_cached_value_after_external_mutation
// origin: languages/kotlin/tests/kotlin/test_delegated_properties.rs

class Cache {
            private var calls = 0
            val value by lazy {
                calls += 1
                calls * 3
            }
            fun currentCalls(): Int = calls
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = Cache()
            __check((c.value).toString(), "3")
            __check((c.value).toString(), "3")
            __check((c.currentCalls()).toString(), "1")
        }
