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
            val c = Cache()
            __p((c.value).toString())
            __p((c.value).toString())
            __p((c.currentCalls()).toString())
        
__check("3\n3\n1")
}
