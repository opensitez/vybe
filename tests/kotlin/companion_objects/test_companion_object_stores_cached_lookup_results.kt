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

        var __buf: String = ""

fun __p(s: String) {
    __buf = __buf + s + "\n"
}

fun __pr(s: String) {
    __buf = __buf + s
}

// The final `println` contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted. Written as two equality
// tests rather than trimming: `String.endsWith` is not implemented in Vybe's
// Kotlin (measured — `"ab\n".endsWith("\n")` throws "undefined is not
// callable"), and a harness that cannot run asserts nothing at all. The cargo
// helper split on "\n" and popped trailing empties, so the two forms were
// equivalent there too.
fun __check(want: String) {
    if (__buf != want && __buf != want + "\n") {
        println("FAIL: want [" + want + "] got [" + __buf + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __p((Dictionary.count()).toString())
            Dictionary.put("x", "one")
            Dictionary.put("y", "two")
            __p((Dictionary.lookup("x")).toString())
            __p((Dictionary.count()).toString())
        
__check("0\none\n2")
}
