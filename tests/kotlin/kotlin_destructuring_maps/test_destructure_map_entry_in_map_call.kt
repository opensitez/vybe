// vybe-test: kotlin/kotlin_destructuring_maps/test_destructure_map_entry_in_map_call
// origin: languages/kotlin/tests/kotlin/test_kotlin_destructuring_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mapOf("a" to 1, "b" to 2)
            val doubled = values.map { (k, v) -> k + v.toString() }
            __check((doubled.joinToString(",")).toString(), "a1,b2")
        }
