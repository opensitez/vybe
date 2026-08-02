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

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val store = StoreProxy(MemoryStore())
            store.put("a", 3)
            __check((store.get("a")).toString(), "3")
        }
