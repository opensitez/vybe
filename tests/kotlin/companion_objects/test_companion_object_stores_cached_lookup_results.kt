// vybe-test: kotlin/companion_objects/test_companion_object_stores_cached_lookup_results
// origin: languages/kotlin/tests/kotlin/test_companion_objects.rs

class Dictionary {
            companion object {
                private val cache = mutableMapOf<String, String>()

                fun put(key: String, value: String) {
                    cache[key] = value
                }

                fun lookup(key: String): String {
                    return cache[key] ?: ""
                }

                fun count(): Int = cache.size
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Dictionary.count()).toString(), "0")
            Dictionary.put("x", "one")
            Dictionary.put("y", "two")
            __check((Dictionary.lookup("x")).toString(), "one")
            __check((Dictionary.count()).toString(), "2")
        }
