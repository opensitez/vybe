// vybe-test: kotlin/equality_hashcode/test_map_with_equivalent_data_key_overwrites_value
// origin: languages/kotlin/tests/kotlin/test_equality_hashcode.rs

data class Key(val id: String)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = hashMapOf(Key("a") to 1)
            map[Key("a")] = 9
            __check((map.size).toString(), "1")
            __check((map[Key("a")]).toString(), "9")
        }
