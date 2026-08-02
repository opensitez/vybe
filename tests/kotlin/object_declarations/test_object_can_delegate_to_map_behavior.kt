// vybe-test: kotlin/object_declarations/test_object_can_delegate_to_map_behavior
// origin: languages/kotlin/tests/kotlin/test_object_declarations.rs

object Cache : Map<String, Int> by mapOf("a" to 1, "b" to 2) {
            val keysText = keys.joinToString("-")
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Cache["a"]).toString(), "1")
            __check((Cache.keysText).toString(), "a-b")
            __check((Cache.size).toString(), "2")
        }
