// vybe-test: kotlin/class_delegation/test_map_delegation_key_lookup
// origin: languages/kotlin/tests/kotlin/test_class_delegation.rs

class ReadOnlyMap(delegate: Map<String, Int>) : Map<String, Int> by delegate

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val m = ReadOnlyMap(mapOf("a" to 1, "b" to 2))
            __check((m["a"]).toString(), "1")
            __check((m.keys.joinToString(",")).toString(), "a,b")
        }
