// vybe-test: kotlin/properties/test_lazy_property_initializes_once_across_multiple_accesses
// origin: languages/kotlin/tests/kotlin/test_properties.rs

class Cache {
            var loads = 0
            val value: String by lazy {
                loads += 1
                "loaded"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val cache = Cache()
            __check((cache.value).toString(), "loaded")
            __check((cache.value).toString(), "loaded")
            __check((cache.loads).toString(), "1")
        }
