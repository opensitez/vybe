// vybe-test: kotlin/equality_hashcode/test_map_lookup_uses_hashcode_and_equals
// origin: languages/kotlin/tests/kotlin/test_equality_hashcode.rs

data class Key(val id: String)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = hashMapOf(Key("a") to 1, Key("b") to 2)
            __check((map[Key("a")]).toString(), "1")
            __check((map[Key("x")]).toString(), "null")
        }
