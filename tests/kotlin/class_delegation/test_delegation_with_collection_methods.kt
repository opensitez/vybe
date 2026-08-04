// vybe-test: kotlin/class_delegation/test_delegation_with_collection_methods
// origin: languages/kotlin/tests/kotlin/test_class_delegation.rs

interface Store {
            fun put(key: String, value: Int)
            fun get(key: String): Int?
        }

        class MemoryStore : Store {
            private val data = mutableMapOf<String, Int>()
            override fun put(key: String, value: Int) { data[key] = value }
            override fun get(key: String): Int? = data[key]
        }

        class StoreProxy(delegate: Store) : Store by delegate

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
            val store = StoreProxy(MemoryStore())
            store.put("a", 3)
            __p((store.get("a")).toString())
        
__check("3")
}
